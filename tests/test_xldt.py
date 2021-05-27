import csv
import os
import unittest
import xldt

ROOT = os.path.dirname(__file__)

class TestDateValue(unittest.TestCase):

    def test_date(self):
        for n in range(-1000000, 1000000):
            y = xldt.year(n)
            m = xldt.month(n)
            d = xldt.day(n)
            self.assertEqual(n, xldt.date(y, m, d), 
                'error at {:04d}-{:02d}-{:02d}'.format(y, m, d))
    
    def test_time(self):
        for n in range(0, 86400):
            t = n / 86400
            h = xldt.hour(t)
            m = xldt.minute(t)
            s = xldt.second(t)
            self.assertEqual(n, round(xldt.time(h, m, s) * 86400), 
                'error at {:02d}:{:02d}:{:02d}'.format(h, m, s))

    def test_date_csv(self):
        with open(os.path.join(ROOT, 'data/date.csv'), newline='') as src:
            reader = csv.DictReader(src, delimiter=';')
            for r in reader:
                y, m, d = map(int, r['DATE'].split('-'))
                v = float(r['VALUE'])
                # If the previous test passed, an error here is bad CSV data.
                self.assertEqual(v, xldt.date(y, m, d), 
                    'bad CSV data at {:04d}-{:02d}-{:02d}'.format(y, m, d))
                w = xldt.weekday(v, int(r['RETURN_TYPE']))
                self.assertEqual(w, int(r['WEEKDAY']),
                    'error at {:04d}-{:02d}-{:02d}'.format(y, m, d))

    def test_week_csv(self):
        with open(os.path.join(ROOT, 'data/week.csv'), newline='') as src:
            reader = csv.DictReader(src, delimiter=';')
            for r in reader:
                y, m, d = map(int, r['DATE'].split('-'))
                t = int(r['RETURN_TYPE'])
                w = int(r['WEEKNUM'])
                self.assertEqual(w, xldt.week(xldt.date(y, m, d), t),
                    'error at {:04d}-{:02d}-{:02d} ({})'.format(y, m, d, t))

    def test_iso_week_csv(self):
        with open(os.path.join(ROOT, 'data/isoweek.csv'), newline='') as src:
            reader = csv.DictReader(src, delimiter=';')
            for r in reader:
                y, m, d = map(int, r['DATE'].split('-'))
                w = int(r['ISOWEEKNUM'])
                self.assertEqual(w, xldt.isoweek(xldt.date(y, m, d)),
                    'error at {:04d}-{:02d}-{:02d}'.format(y, m, d))

    def test_date_delta_csv(self):
        with open(os.path.join(ROOT, 'data/delta.csv'), newline='') as src:
            reader = csv.DictReader(src, delimiter=';')
            for r in reader:
                y1, m1, d1 = map(int, r['START_DATE'].split('-'))
                y2, m2, d2 = map(int, r['END_DATE'].split('-'))
                x = int(r['YEARS'])
                self.assertEqual(x, xldt.years(xldt.date(y1, m1, d1),
                                               xldt.date(y2, m2, d2)),
                    'error at {:04d}-{:02d}-{:02d}'.format(y1, m1, d1))
                x = int(r['MONTHS'])
                self.assertEqual(x, xldt.months(xldt.date(y1, m1, d1),
                                               xldt.date(y2, m2, d2)),
                    'error at {:04d}-{:02d}-{:02d}'.format(y1, m1, d1))
                x = int(r['DAYS'])
                self.assertEqual(x, xldt.days(xldt.date(y1, m1, d1),
                                               xldt.date(y2, m2, d2)),
                    'error at {:04d}-{:02d}-{:02d}'.format(y1, m1, d1))

if __name__ == '__main__':
    unittest.main()
