import unittest
import pandas as pd
import numpy as np
from pandas.util.testing import assert_frame_equal

def vlookup(left, right, left_key, right_key, right_col):
    mleft = left.loc[:, left_key].to_frame()
    mright = right.loc[:, [right_key, right_col]]
    output = pd.merge(mleft, mright, how='left', left_on=left_key, right_on=right_key)
    return output.loc[:, right_col].to_frame()

def vlookup_update(left, right, left_key, right_key, left_col, right_col):
    mleft = left.loc[:, [left_key, left_col]]
    mright = right.loc[:, [right_key, right_col]]
    mleft.columns, mright.columns = ['key','lval'], ['key','rval']
    output = pd.merge(mleft, mright, how='left', on='key')
    output.loc[output.loc[:, 'rval'].notnull(), 'lval'] = output.loc[:, 'rval']
    return output.loc[:, 'lval'].to_frame()

class TestStringMethods(unittest.TestCase):

	def setUp(self):
		self.df1 = pd.DataFrame({'a':[1,2], 
			                     'b':[3,4]})
		self.df2 = pd.DataFrame({'c':[3,4],
			                     'd':[7,8]})
		self.df3 = pd.DataFrame({'c':[3,4],
			                     'd':[7,8],
			                     'e':[9,10],
			                     'f':[11,12]})
		self.df4 = pd.DataFrame({'a':[0,1,2,3], 
			                     'b':[99,-1,-1,99]})

	def tearDown(self):
		self.df1 = None
		self.df2 = None

	def test_replace_text_1(self):
		self.assertEqual('Non-Exempt', 'Nonexempt'.replace("N|Nonexempt", "Non-Exempt"))
		self.assertEqual('Non-Exempt', 'N'.replace("N|Nonexempt", "Non-Exempt"))
		self.assertEqual('Exempt', 'Exempt'.replace("Y|Exempt", "Exempt"))
		self.assertEqual('Exempt', 'Y'.replace("Y|Exempt", "Exempt"))

	def test_vlookup_update_1(self):
		actual = self.df4.copy()
		actual['b'] = vlookup_update(self.df4, self.df1, 'a', 'a', 'b', 'b')
		actual['b'] = actual['b'].astype(int) # merge casts 'b' to float -> need to cast back
		expected = pd.DataFrame({'a':[0,1,2,3], 
			                     'b':[99,3,4,99]})
		assert_frame_equal(expected, actual, check_names=False)

	def test_vlookup_1(self):
		self.df1.loc[:, 'd'] = vlookup(self.df1, self.df2, 'b', 'c', 'd')
		actual = vlookup(self.df1, self.df2, 'b', 'c', 'd')
		expected = pd.DataFrame({'d':[7,8]})
		assert_frame_equal(expected, actual, check_names=False)

	def test_vlookup_2(self):
		actual = self.df1.copy()
		actual['d'] = vlookup(actual, self.df2, 'b', 'c', 'd')
		expected = pd.DataFrame({'a':[1,2],
			                     'b':[3,4],
			                     'd':[7,8]})
		assert_frame_equal(expected, actual, check_names=True)

	def test_vlookup_3(self):
		actual = self.df1.copy()
		actual.loc[:,'d'] = vlookup(actual, self.df3, 'b', 'c', 'd')
		actual.loc[:,'e'] = vlookup(actual, self.df3, 'b', 'c', 'e')
		actual.loc[:,'f'] = vlookup(actual, self.df3, 'b', 'c', 'f')
		expected = pd.DataFrame({'a':[1,2],
			                     'b':[3,4],
			                     'd':[7,8],
			                     'e':[9,10],
			                     'f':[11,12]})
		assert_frame_equal(expected, actual, check_names=True)

	def test_format_number_1(self):
		self.assertEqual('000001', '%06d' % 1)
		self.assertEqual('000012', '%06d' % 12)
		self.assertEqual('000123', '%06d' % 123)
		self.assertEqual('001234', '%06d' % 1234)
		self.assertEqual('012345', '%06d' % 12345)
		self.assertEqual('123456', '%06d' % 123456)

	def test_format_number_2(self):
		actual = pd.DataFrame({'a':[1,12,123,1234,12345,123456]})
		actual['a'] = actual['a'].apply('{0:0>6}'.format)
		expected = pd.DataFrame({'a':['000001','000012','000123','001234','012345','123456']})
		assert_frame_equal(expected, actual, check_names=True)

	def test_format_number_3(self):
		actual = pd.DataFrame({'a':['1','12','123','1234','12345','123456']})
		actual['a'] = actual['a'].apply('{0:0>6}'.format)
		expected = pd.DataFrame({'a':['000001','000012','000123','001234','012345','123456']})
		assert_frame_equal(expected, actual, check_names=True)

	def test_format_number_4(self):
		actual = pd.DataFrame({'a':['A1','A12','A123','A1234','A123456','A1234567','A12345678','A123456789']})
		actual['a'] = actual['a'].apply('{0:0>6}'.format)
		expected = pd.DataFrame({'a':['0000A1','000A12','00A123','0A1234','A123456','A1234567','A12345678','A123456789']})
		assert_frame_equal(expected, actual, check_names=True)

if __name__ == '__main__':
	unittest.main()
















