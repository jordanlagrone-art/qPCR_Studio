import base64
import io
import math
import unittest
from PIL import Image
from engine import analyze, statistical_test, render


def fixture():
    raw=[]
    for t in range(3):
        for condition, values in [('Control',[20,25,25]),('Treatment',[20,23+t*.2,23+t*.2]),('Missing',[None,24,24])]:
            for i,(target,ct) in enumerate(zip(['GAPDH','GENE','GENE'],values)):
                if ct is not None:raw.append(dict(trialId=str(t),sample=condition,target=target,cq=ct,omit=False))
    return dict(raw=raw,trials=[dict(id=str(t),name=f'Trial {t+1}') for t in range(3)],sampleMap={},selectedGenes=['GENE'],refGene='GAPDH',calibrator='Control',base=2)


class EngineTests(unittest.TestCase):
    def test_normalization_and_missing_reference(self):
        p=fixture();res=analyze(p)['results']
        self.assertEqual(res[1]['expression'],4)
        self.assertEqual(res[1]['nTarget'],2)
        self.assertEqual(res[2]['expression'],None)
        self.assertEqual(res[0]['expression'],1)
        p['base']=1.9
        self.assertAlmostEqual(analyze(p)['results'][1]['expression'],1.9**2)
        p['raw'][4]['omit']=True
        self.assertEqual(analyze(p)['results'][1]['nTarget'],1)

    def test_paired_matches_trial_ids(self):
        rows=[dict(trialId=str(i),condition=c,value=v) for c,vs in [('a',[1,2,3]),('b',[2,4,6])] for i,v in enumerate(vs)]
        result=statistical_test(dict(rows=list(reversed(rows)),test='paired'))
        self.assertAlmostEqual(result['p'],0.07417990022744853)
        self.assertIn('n = 3',result['detail'])
        with self.assertRaises(ValueError):statistical_test(dict(rows=rows[:1]+rows[3:4],test='paired'))

    def test_welch_and_anova(self):
        rows=[dict(trialId=str(i),condition=c,value=v) for c,vs in [('a',[1,2,3]),('b',[2,4,6])] for i,v in enumerate(vs)]
        self.assertAlmostEqual(statistical_test(dict(rows=rows,test='welch'))['p'],0.2208808404940958)
        result=statistical_test(dict(rows=rows,test='anova'))
        self.assertAlmostEqual(result['statistic'],2.4)
        self.assertAlmostEqual(result['p'],1-math.sqrt(.375)*(1.5-.5*.375))

    def test_figures_and_formats(self):
        rows=[dict(r,value=r['expression']) for r in analyze(fixture())['results']]
        payload=dict(rows=rows,genes=['GENE'],colors=['#48735d','#ba8e65'],width=640,controls=dict(graphTitle='Expression',showGrid=False))
        for kind in ['bar','grouped','points','paired','box']:
            payload['controls']['graphType']=kind
            res=render(payload)
            self.assertEqual(Image.open(io.BytesIO(base64.b64decode(res['data']))).size,(768,576))
        for fmt in ['svg','pdf']:
            res=render(dict(payload,format=fmt,export=True));data=base64.b64decode(res['data'])
            self.assertIn(b'<svg' if fmt=='svg' else b'%PDF',data)
        payload['controls']['yScale']='log';payload['rows'][0]['value']=0
        self.assertEqual(render(payload)['omitted'],1)


if __name__=='__main__':unittest.main()
