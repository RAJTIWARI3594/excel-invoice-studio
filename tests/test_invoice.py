import unittest,tempfile,json
from pathlib import Path
from openpyxl import load_workbook
from invoice import validate,totals,build,money
class Tests(unittest.TestCase):
 def setUp(self):self.data=json.loads(Path('sample/invoice.json').read_text())
 def test_totals(self):self.assertEqual(totals(self.data),{'net':2270.0,'tax':218.36,'total':2488.36,'balance':1988.36})
 def test_round(self):self.data['lines']=[{'description':'x','quantity':1,'unit_price':'0.10','tax_percent':'5'}];self.data['paid']=0;self.assertEqual(totals(self.data)['tax'],.01)
 def test_quantity(self):
  self.data['lines'][0]['quantity']=True
  with self.assertRaises(ValueError):validate(self.data)
 def test_negative(self):
  with self.assertRaises(ValueError):money('-1')
 def test_precision(self):
  with self.assertRaises(ValueError):money('1.001')
 def test_nan(self):
  with self.assertRaises(ValueError):money('NaN')
 def test_tax(self):
  self.data['lines'][0]['tax_percent']=101
  with self.assertRaises(ValueError):validate(self.data)
 def test_formula_and_literal(self):
  self.data['customer']='=1+1'
  with tempfile.TemporaryDirectory() as d:
   p=Path(d)/'x.xlsx';r=build(self.data,p);s=load_workbook(p).active;self.assertEqual(s['B5'].data_type,'s');self.assertEqual(s['E8'].value,'=ROUND(B8*C8,2)');self.assertEqual(r['total_cell'],'G14');self.assertEqual(len(s.data_validations.dataValidation),3)
 def test_overwrite(self):
  with tempfile.TemporaryDirectory() as d:
   p=Path(d)/'x.xlsx';build(self.data,p)
   with self.assertRaises(ValueError):build(self.data,p)
 def test_empty(self):
  self.data['lines']=[]
  with self.assertRaises(ValueError):validate(self.data)
if __name__=='__main__':unittest.main()
