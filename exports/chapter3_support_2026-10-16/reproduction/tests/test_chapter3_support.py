import csv
import hashlib
import json
from decimal import Decimal as D
from pathlib import Path
import unittest

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'outputs/chapter3_support'


def rows(name):
    with (OUT/(name+'.csv')).open(encoding='utf-8-sig',newline='') as f:
        return list(csv.DictReader(f))


class Chapter3SupportTests(unittest.TestCase):
    def test_contract_amounts_do_not_become_actual_cash(self):
        payments=rows('项目付款条件')
        self.assertTrue(all(r['actual_cash_received_cny']=='' for r in payments))
        self.assertEqual(sum(D(r['nominal_original_price_calculation_cny']) for r in payments if r['source_ids']=='TT'),D('96830100'))
        self.assertTrue(all(r['nominal_original_price_calculation_cny']=='' for r in payments if r['source_ids']!='TT'))

    def test_price_bridge_conflicts_preserved(self):
        price=rows('项目价款调整')
        self.assertEqual([D(r['unreconciled_difference_cny']) for r in price],[D('321305'),D('12905050')])
        self.assertEqual(sum(D(r['reduction_cny']) for r in price),D('214385700'))
        self.assertIn('D02',{r['id'] for r in rows('争议与不可推断项')})

    def test_subsequent_balances_not_imputed_receipts(self):
        r=rows('期后回收观察')
        self.assertEqual(D(r[0]['net_decrease_cny']),D('39400000'))
        self.assertTrue(all(x['actual_receipt_by_counterparty_cny']=='' for x in r))
        self.assertTrue(all('未经审计' in x['status'] for x in r))
        self.assertEqual({int(x['year']) for x in rows('描述性财务指标')},set(range(2019,2026)))

    def test_scope_and_missing_statistics_preserved(self):
        r=rows('行业装机统计')
        self.assertEqual(next(x for x in r if x['year']=='2023')['new_capacity_10k_kW'],'')
        self.assertTrue(all('非吊装' in x['definition'] for x in r))
        sources=rows('来源目录')
        self.assertTrue(all('转载' in x['source_type'] for x in sources if x['source_id'] in ['J24','J25']))

    def test_manifest_fingerprints_and_original_quotes(self):
        manifest=json.loads((OUT/'run_manifest.json').read_text())
        self.assertEqual(manifest['script_sha256'],hashlib.sha256((ROOT/'scripts/build_chapter3_support_pack.py').read_bytes()).hexdigest())
        sources={r['source_id']:r for r in manifest['sources']}
        for source in sources.values():
            self.assertEqual(source['sha256'],hashlib.sha256((ROOT/source['local_path']).read_bytes()).hexdigest())
        import re
        compact=lambda x:re.sub(r'\s+','',x)
        for r in rows('原文证据索引'):
            self.assertIn(compact(r['quote']),compact((ROOT/sources[r['source_id']]['text_path']).read_text()))


if __name__=='__main__':unittest.main()
