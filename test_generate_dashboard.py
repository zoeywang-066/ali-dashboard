import io
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import generate_dashboard as dashboard


class Source:
    def __init__(self, name, modified):
        self.name = name
        self.modified = modified

    def exists(self):
        return True

    def stat(self):
        return SimpleNamespace(st_mtime=self.modified)


class DataSourceTests(unittest.TestCase):
    def load_sources(self, master_modified, detail_modified):
        master = Source('master.xlsx', master_modified)
        detail = Source('detail.xlsx', detail_modified)
        datasets = {
            'master.xlsx': [{'ds':'20261122', 'cid':'a', 'name':'AliExpress_new', 'gmv':20}],
            'detail.xlsx': [
                {'ds':'20261122', 'cid':'a', 'name':'AliExpress_old', 'gmv':10},
                {'ds':'20261121', 'cid':'a', 'name':'AliExpress_old', 'gmv':5},
            ],
        }
        with patch.multiple(
            dashboard,
            XLSX_FILE=master,
            DOWNLOADS=SimpleNamespace(glob=lambda pattern:[detail]),
            HISTORY_FILE=SimpleNamespace(exists=lambda:False),
            _load_xlsx=lambda source:[dict(row) for row in datasets[source.name]],
            _apply_moloco_local_spend=lambda rows:rows,
        ):
            return dashboard.load_data()

    def test_new_master_overrides_old_details(self):
        rows = self.load_sources(2, 1)
        latest = [row for row in rows if row['ds']=='20261122']
        self.assertEqual(len(latest), 1)
        self.assertEqual(latest[0]['gmv'], 20)
        self.assertTrue(any(row['ds']=='20261121' for row in rows))

    def test_new_detail_overrides_old_master(self):
        rows = self.load_sources(1, 2)
        self.assertEqual(next(row for row in rows if row['ds']=='20261122')['gmv'], 10)

    def test_explicit_rebate_zero_negative_and_missing(self):
        records = [
            {'ds':'20261122', 'cid':cid, 'name':'AliExpress_rmkt',
             'source_sheet':'最新', 'spend':100, 'gmv':200, 'rebate_spend':rebate}
            for cid,rebate in [('a',80), ('b',0), ('c',-5), ('d',None)]
        ]
        cached = 'date_local,campaign_id,campaign_name,country,spend_usd\n'
        cached += ''.join(f'2026-11-22,{cid},AliExpress_rmkt,GB,120\n' for cid in 'abcd')
        with patch.object(Path, 'open', return_value=io.StringIO(cached)):
            result = dashboard._apply_moloco_local_spend(records, 'test.csv')
        self.assertEqual([row['rebate_spend'] for row in result], [80,0,-5,100])
        self.assertEqual(sum(row['spend'] for row in result), 480)
        self.assertEqual(result[0]['rebate_roi'], 2.5)
        self.assertIsNone(result[1]['rebate_roi'])


if __name__ == '__main__':
    unittest.main()
