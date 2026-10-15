"""Integration checks for the actual catalogue and returned Excel copies."""
import json
from pathlib import Path
import shutil
import tempfile
import unittest

from openpyxl import load_workbook
import catalog_tools as m


class CatalogueTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data=m.catalog()
        cls.template=m.ROOT/m.SURVEY
        cls.shared=next(t for t in cls.data['topics'] if len(m.codes(t))>=3)
        cls.single=next(t for t in cls.data['topics'] if len(m.codes(t))==1)
        cls.other=next(t for t in cls.data['topics'] if t['id'] not in [cls.shared['id'],cls.single['id']])
        cls.row_by_id={t['id']:i for i,t in enumerate(cls.data['topics'],12)}

    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.folder=Path(self.temp.name)

    def tearDown(self):self.temp.cleanup()

    def response(self,filename,name,contact,answers):
        w=load_workbook(self.template)
        s=w['Выбор тем'];s['B4']=name;s['F5']=contact;s['B6']='Выбор завершён'
        for topic,ready,scope,comment in answers:
            row=self.row_by_id[topic['id']]
            s.cell(row,5).value=ready;s.cell(row,6).value=scope;s.cell(row,7).value=comment
        path=self.folder/filename;w.save(path);w.close();return path

    def test_source_coverage_and_blank_template(self):
        d=self.data
        self.assertEqual(len(d['programs']),18)
        self.assertEqual(sum(len(t['usages']) for t in d['topics']),329)
        self.assertEqual(len(d['topics']),265)
        self.assertEqual(len({t['id'] for t in d['topics']}),265)
        self.assertEqual(sum(not u['content'] for t in d['topics'] for u in t['usages']),3)
        w=load_workbook(self.template)
        self.assertIsNone(w['Выбор тем']['B4'].value)
        self.assertIsNone(w['Выбор тем']['F5'].value)
        for row,t in enumerate(d['topics'],12):
            self.assertEqual(w['Выбор тем'].cell(row,1).value,t['id'])
            self.assertIsNone(w['Выбор тем'].cell(row,5).value)
            self.assertIsNone(w['Выбор тем'].cell(row,7).value)
            target=w['Выбор тем'].cell(row,4).hyperlink.location.split('!')[1]
            self.assertEqual(w['Содержание тем'][target].value,t['title'])
        self.assertFalse(any(c.comment for s in w for row in s for c in row))
        self.assertFalse(w._external_links)
        w.close()

    def test_program_scope_and_partial_readiness(self):
        selected=m.codes(self.shared)[0]
        a=self.response('a.xlsx','Тестовый лектор 1','test1@example.invalid',[
            (self.shared,'Да',selected,''),(self.single,'Частично','Все','Только теория')])
        b=self.response('b.xlsx','Тестовый лектор 2','test2@example.invalid',[
            (self.shared,'Да','Все',''),(self.other,'Да','Все','')])
        result=m.make_registry(self.data,[a,b],self.folder/'svod.xlsx')
        self.assertEqual(result['forms'],2);self.assertEqual(result['marks'],4)
        expected=1+len(m.codes(self.single))+len(m.codes(self.shared))+len(m.codes(self.other))
        self.assertEqual(result['topic_program_lecturer_links'],expected)
        w=load_workbook(self.folder/'svod.xlsx',data_only=True)
        rows=list(w['Реестр'].iter_rows(min_row=6,values_only=True))
        lecturer1=[r for r in rows if r[2]=='Тестовый лектор 1' and r[8]==self.shared['id']]
        self.assertEqual([r[1] for r in lecturer1],[selected])
        partial=[r for r in rows if r[2]=='Тестовый лектор 1' and r[3]=='Частично']
        self.assertTrue(partial);self.assertTrue(all(r[7]=='Уточнить объём' for r in partial))
        self.assertTrue(any(not r[2] and r[7]=='Нет отметок' for r in rows))
        w.close()

    def test_multiple_codes_and_sorting(self):
        selected=m.codes(self.shared)[:2]
        p=self.response('a.xlsx','Тест','id-1',[(self.shared,'Да','; '.join(selected),'')])
        original=m.read_response(p,self.data)
        self.assertEqual(original['answers'][0]['programs'],selected)
        w=load_workbook(p);s=w['Выбор тем'];last=11+len(self.data['topics'])
        values=[[s.cell(r,c).value for c in range(1,8)] for r in range(12,last+1)]
        for r,row in zip(range(12,last+1),reversed(values)):
            for c,value in enumerate(row,1):s.cell(r,c).value=value
        w.save(p);w.close()
        self.assertEqual(m.read_response(p,self.data)['answers'],original['answers'])

    def test_duplicate_and_conflicting_versions(self):
        a=self.response('a.xlsx','Тест','id-1',[(self.shared,'Да','Все','')])
        b=self.folder/'duplicate.xlsx';shutil.copyfile(a,b)
        people,duplicates=m.responses(self.data,[a,b])
        self.assertEqual(len(people),1);self.assertEqual(len(duplicates),1)
        w=load_workbook(b);s=w['Выбор тем'];s.cell(self.row_by_id[self.shared['id']],5).value='Частично';w.save(b);w.close()
        with self.assertRaises(ValueError):m.make_registry(self.data,[a,b],self.folder/'bad.xlsx')
        self.assertFalse((self.folder/'bad.xlsx').exists())

    def test_invalid_responses(self):
        p=self.response('a.xlsx','Тест','id-1',[(self.shared,'Да','Все','')])
        row=self.row_by_id[self.shared['id']]
        for cell,value in [('B4',''),('B5','WRONG-VERSION'),(f'F{row}','НЕИЗВЕСТНО'),(f'E{row}','=1+1'),(f'A{row}','T-UNKNOWN')]:
            w=load_workbook(p);w['Выбор тем'][cell]=value;q=self.folder/'bad.xlsx';w.save(q);w.close()
            with self.assertRaises(ValueError,msg=cell):m.read_response(q,self.data)

    def test_empty_response_and_homonyms(self):
        a=self.response('a.xlsx','Одинаковое ФИО','id-1',[])
        b=self.response('b.xlsx','Одинаковое ФИО','id-2',[])
        people,_=m.responses(self.data,[a,b]);self.assertEqual(len(people),2)
        result=m.make_registry(self.data,[a],self.folder/'empty.xlsx')
        self.assertEqual(result['marks'],0)
        w=load_workbook(self.folder/'empty.xlsx');self.assertEqual(w['Реестр'].max_row-5,329);w.close()
        c=self.response('c.xlsx','Одинаковое ФИО','',[])
        with self.assertRaises(ValueError):m.responses(self.data,[a,c])


if __name__=='__main__':unittest.main()
