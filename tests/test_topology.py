import copy
import importlib.util
import json
import random
import tempfile
import unittest
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
module_spec=importlib.util.spec_from_file_location('topology',ROOT/'skills/visio-network-topology/scripts/topology.py')
t=importlib.util.module_from_spec(module_spec)
module_spec.loader.exec_module(t)


class TopologyTests(unittest.TestCase):
    def setUp(self):
        self.rows,self.source=t.read_rows(ROOT/'examples/links.csv')

    def test_complete_dual_core_graph_and_labels(self):
        model=t.plan(self.rows,self.source)
        report=t.validate_plan(model)
        self.assertEqual((report['devices'],report['links'],report['port_labels']),(14,28,56))
        self.assertEqual(report['line_types'],{'40G':2,'10G':24,'copper':2})
        for role in ('firewall','core','access'):
            self.assertEqual(len({n['cy'] for n in model['nodes'] if n['role']==role}),1)
        fw=[n for n in model['nodes'] if n['role']=='firewall']
        self.assertAlmostEqual(fw[1]['cx']-fw[0]['cx'],3.4)

    def test_renaming_and_shuffling_do_not_require_example_names(self):
        rows=copy.deepcopy(self.rows)
        names=sorted({r[k] for r in rows for k in ('from','to')})
        renamed={name:f'设备 {i+1}' for i,name in enumerate(names)}
        for row in rows:
            for key in ('from','to'):
                row[key]=renamed[row[key]]
        random.Random(11).shuffle(rows)
        model=t.plan(rows,self.source)
        self.assertEqual(t.validate_plan(model)['port_labels'],56)
        self.assertEqual({n['name'] for n in model['nodes']},set(renamed.values()))

    def test_reversed_duplicate_preserves_parallel_physical_links(self):
        row=self.rows[0]
        reverse={k:row[k] for k in t.FIELDS}
        for suffix in ('','_type','_port'):
            reverse['from'+suffix],reverse['to'+suffix]=row['to'+suffix],row['from'+suffix]
        nodes,links,duplicates=t.normalize(self.rows+[reverse],{})
        self.assertEqual(len(links),28)
        self.assertEqual(len(duplicates),1)
        self.assertEqual(sum(l['category']=='40G' for l in links),2)

    def test_port_reuse_with_different_peer_is_rejected(self):
        rows=copy.deepcopy(self.rows)
        rows[3]['from_port']=rows[2]['from_port']
        with self.assertRaisesRegex(t.TopologyError,'port reused'):
            t.normalize(rows,{})

    def test_unknown_type_missing_field_and_color_are_rejected(self):
        for field,value in [('from_type','Mystery'),('from_port',''),('line_type','25G')]:
            rows=copy.deepcopy(self.rows)
            rows[0][field]=value
            with self.subTest(field=field),self.assertRaises(t.TopologyError):
                t.normalize(rows,{})

    def test_column_order_is_not_fixed(self):
        output_root=(ROOT/'outputs').resolve()
        output_root.mkdir(exist_ok=True)
        with tempfile.TemporaryDirectory(dir=output_root) as directory:
            self.assertTrue(Path(directory).resolve().is_relative_to(output_root))
            path=Path(directory)/'links.csv'
            import csv
            with path.open('w',encoding='utf-8',newline='') as stream:
                writer=csv.writer(stream)
                writer.writerow(list(reversed(t.HEADERS)))
                writer.writerow(list(reversed([self.rows[0][k] for k in t.FIELDS])))
            rows,_=t.read_rows(path)
            self.assertEqual(rows[0]['from_port'],'E1/47')

    def test_verifier_rejects_label_drifting_off_its_wire(self):
        model=t.plan(self.rows,self.source)
        tag=model['labels'][0]
        tag['y']+=.4
        with self.assertRaisesRegex(t.TopologyError,'inline port'):
            t.validate_plan(model)

    def test_type_extension_and_color_mapping(self):
        rows=copy.deepcopy(self.rows)
        for row in rows:
            for side in ('from','to'):
                if row[side+'_type']=='Core-switch':
                    row[side+'_type']='主干'
            if row['line_type']=='40G光纤':
                row['line_type']='40G DAC'
        config={'device_types':{'主干':{'role':'core','master':'Core-Switch'}},'line_types':{'40G DAC':'40G'}}
        self.assertEqual(t.validate_plan(t.plan(rows,self.source,config))['links'],28)


if __name__=='__main__':
    unittest.main()
