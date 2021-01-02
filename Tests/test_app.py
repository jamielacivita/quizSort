import unittest
from App import foo

class test_load_roster_data(unittest.TestCase):
	def test_load_roster_data(self):
		result = foo.load_roster_data()
		self.assertEqual(result, "load roster data")


class test_load_quiz_data(unittest.TestCase):
	def test_load_quiz_data(self):
		result = foo.load_quiz_data()
		self.assertEqual(result, "load quiz data")



class test_load_data_student_score_class(unittest.TestCase):
	def test_data_student_score_class(self):
		result = foo.load_data_student_score_class()
		self.assertEqual(result, "load data student score class")


class test_generate_csv_files(unittest.TestCase):
	def test_generate_csv_files(self):
		result = foo.generate_csv_files()
		self.assertEqual(result, "generate csv files")


class test_read_cell_A5(unittest.TestCase):
	def test_read_cell_A5(self):
		result = foo.readCellA5()
		self.assertEqual(result, "emma mcgreevy* ( emma mcgreevy )")


class test_read_cell_B2(unittest.TestCase):
	def test_read_cell_B2(self):
		result = foo.readCellB2()
		self.assertEqual(result, "MADISON ALBRECHT")


class test_get_first_class_students_list(unittest.TestCase):
	def test_first_class_students_list(self):
		results_lst = foo.getFirstClassStudents_lst()
		self.assertEqual(results_lst[0], "MADISON ALBRECHT")
		self.assertEqual(results_lst[1], "MAKAYLA ALLEN")
		self.assertEqual(results_lst[2], "KARINE CAJIGAS")


class test_parse_player_levelR5(unittest.TestCase):
	def test_parse_player_levelR5(self):
		results_tpl = foo.parsePlayerLevelR5()
		self.assertEqual(results_tpl[0], "EMMA MCGREEVY")
		self.assertEqual(results_tpl[1], "18600")
		self.assertEqual(results_tpl[2], "100%")


class test_parse_player_level(unittest.TestCase):
	def test_parse_player_level_row5(self):
		results_tpl = foo.getParsePlayerLevelRow(5)
		self.assertEqual(results_tpl[0], "EMMA MCGREEVY")
		self.assertEqual(results_tpl[1], "18600")
		self.assertEqual(results_tpl[2], "100%")
	def test_parse_player_level_row6(self):
		results_tpl = foo.getParsePlayerLevelRow(6)
		self.assertEqual(results_tpl[0], "FAIZA ROGARIA")
		self.assertEqual(results_tpl[1], "18000")
		self.assertEqual(results_tpl[2], "100%")
	def test_parse_player_level_row90(self):
		results_tpl = foo.getParsePlayerLevelRow(90)
		self.assertEqual(results_tpl[0], "ISABELLA LOPEZ")
		self.assertEqual(results_tpl[1], "0")
		self.assertEqual(results_tpl[2], "0%")


class test_parse_player(unittest.TestCase):
	def test_parse_player(self):
		player_level_dict = foo.parsePlayer()
		emma = player_level_dict["EMMA MCGREEVY"]
		self.assertEqual(emma[0] , "EMMA MCGREEVY")
		self.assertEqual(emma[1], "18600")
		self.assertEqual(emma[2], "100%")



if __name__ == "__main__":
	unittest.main()
