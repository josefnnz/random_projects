import unittest
import pandas as pd
import numpy as np
from pandas.util.testing import assert_frame_equal

def vlookup(left, right, left_key, right_key, right_col):
    mleft = left.loc[:, left_key].to_frame()
    mright = right.loc[:, [right_key, right_col]]
    output = pd.merge(mleft, mright, how='left', left_on=left_key, right_on=right_key)
    return output.loc[:, right_col].to_frame()

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

	def tearDown(self):
		self.df1 = None
		self.df2 = None

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

if __name__ == '__main__':
	unittest.main()