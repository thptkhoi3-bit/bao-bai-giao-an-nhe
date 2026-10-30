import sys,json,docx
sys.path.insert(0,'.')
import core
A='/home/box/agent-data/agents/559020c0-3f09-44aa-b19e-a28878d0460c/attachments/'
r=core.fill_baobai(A+'5b1a0fb26dd5dd40f6dccae287964e5ef8f1dd0965ffc9d173ade3627b5f9d36.docx',A+'455b7952f7261379a94cf4da87ee94402e99a80bdba6eb6cdc2e04286f978ff2.docx','/tmp/t_bb.docx',7)
print('weeks',[w['week'] for w in r['weeks']],'skipped',r['skipped_weeks'],'empty',r['weeks'][0]['empty'],'unused',r['weeks'][0]['unused'])
a=docx.Document('/tmp/t_bb.docx');b=docx.Document('/cursor/stores/store-0e61161c-f9bc-45cb-8d83-359add90b33e/baobai/BaoBai_Goc_moinhat.docx')
print('BAOBAI identical text:',[[c.text for r in t.rows for c in r.cells] for t in a.tables]==[[c.text for r in t.rows for c in r.cells] for t in b.tables])
r2=core.build_giaoan('/workspace/GiaoAn_Goc.docx','/tmp/t_bb.docx',7,[A+'5cc1bcb24bf4feb4d1794c9efb553937fe7421deb25a4b96bcc6c150caaf2e1b.docx'],'/tmp/t_ga.docx')
for t in r2['tiet']: print(t['thu'],t['tiet'][:30],'|',t['inserted'])
print('unused',r2['unused']);print('step1',r2['step1']);print(r2['dates'])
from docx.oxml.ns import qn
tx=lambda d:[''.join(x.text or '' for x in e.iter(qn('w:t'))) for e in docx.Document(d).element.body]
X,Y=tx('/tmp/t_ga.docx'),tx('/workspace/GiaoAn_Tuan7_HoanChinh_v2.docx')
print('GIAOAN identical text:',X==Y,len(X),len(Y))
if X!=Y:
  import difflib
  for l in list(difflib.unified_diff(Y,X,lineterm='',n=0))[:20]: print(l[:120])
