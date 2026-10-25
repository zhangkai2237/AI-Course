import json,unittest
from unittest.mock import patch
import server
class BackendTests(unittest.TestCase):
 def setUp(self):
  self.source={'id':'F-PROMO-01','version':1,'sku':'P1标准版','activity':'P1十月活动','from':'2026-10-01','to':'2026-10-31','enabled':True,'text':'原价2399元，新会员领用400元活动券及200元专享券后1799元。'}
  self.snapshot={'task':{'text':'本次活动1799元，人人同价。','sku':'P1标准版','activity':'P1十月活动','asOf':'2026-10-21'},'config':{'rules':[{'id':'R02','enabled':True,'source':'F-PROMO-01','instruction':'结合全文核对条件'}],'sources':[self.source]}}
  self.raw={'checks':[{'ruleId':'R02','status':'issue','reason':'缺少资格及用券条件','quote':'本次活动1799元，人人同价。','sourceId':'F-PROMO-01','sourceVersion':1,'evidence':self.source['text'],'advice':'补齐条件'}]}
 def test_valid_response(self):
  p,b=server.prepare_input(self.snapshot);rows,err=server.validate_output(self.raw,p,b);self.assertEqual(rows[0]['status'],'issue');self.assertFalse(err)
 def test_fake_quote_rejected(self):
  self.raw['checks'][0]['quote']='不存在的原句';p,b=server.prepare_input(self.snapshot);rows,_=server.validate_output(self.raw,p,b);self.assertEqual(rows[0]['status'],'incomplete')
 def test_wrong_version_rejected(self):
  self.raw['checks'][0]['sourceVersion']=2;p,b=server.prepare_input(self.snapshot);rows,_=server.validate_output(self.raw,p,b);self.assertEqual(rows[0]['status'],'incomplete')
 def test_missing_material_blocked(self):
  self.source['activity']='另一活动';p,b=server.prepare_input(self.snapshot);self.assertFalse(p['rules']);self.assertEqual(b[0]['status'],'incomplete')
 def test_duplicate_or_missing_check_not_pass(self):
  p,b=server.prepare_input(self.snapshot);self.raw['checks']*=2;rows,_=server.validate_output(self.raw,p,b);self.assertEqual(rows[0]['status'],'incomplete')
 def test_plain_http_remote_rejected(self):
  with self.assertRaises(ValueError):server.validate_base('http://example.com/v1')
 def test_credentials_not_exposed(self):
  c=server.safe_config({'base':'https://example.invalid/v1','model':'test-model','key':'test-secret'});self.assertNotIn('test-secret',json.dumps(c));self.assertTrue(c['configured'])
 def test_live_adapter_with_mock_provider(self):
  raw=self.raw
  class Response:
   def __enter__(self):return self
   def __exit__(self,*a):pass
   def read(self,n):return json.dumps({'choices':[{'message':{'content':json.dumps(raw)}}],'usage':{'total_tokens':50}}).encode()
  class Opener:
   def open(self,req,timeout):
    body=json.loads(req.data);assert body['messages'][0]['role']=='system';assert body['model']=='test-model';assert 'snapshot' not in body
    return Response()
  with patch.object(server,'build_opener',return_value=Opener()):r=server.call_model(self.snapshot,{'base':'https://example.invalid/v1','model':'test-model','key':'test-secret'})
  self.assertEqual(r['engine'],'live');self.assertEqual(r['checks'][0]['status'],'issue');self.assertTrue(r['networkCalled'])
 def test_bad_model_json_is_incomplete_error(self):
  with self.assertRaises(ValueError):server.validate_output({'oops':[]},{'task':self.snapshot['task'],'rules':self.snapshot['config']['rules'],'sources':[self.source]},[])
if __name__=='__main__':unittest.main(verbosity=2)
