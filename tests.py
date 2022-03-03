import sys
sys.path.append('.')
import unittest
import pprint
pp = pprint.PrettyPrinter(indent=2, width=1)
import main
import io


class TestXlTex(unittest.TestCase):
  def setUp(self):
    pass

  def test_gera_marcador_invertido_fechamento(self):
    res = main.gera_marcador_invertido(' {{xx')
    #print(res)
    self.assertEqual(res, 'xx}}')

  def test_gera_marcador_invertido_abertura(self):
    res = main.gera_marcador_invertido('xx}} ')
    self.assertEqual(res, '{{xx')

  def test_gera_marcador_invertido_fechamento_extras(self):
    res = main.gera_marcador_invertido(' {{xx   asd   ')
    self.assertEqual(res, 'xx}}')

  def test_gera_marcador_invertido_abertura_extras(self):
    res = main.gera_marcador_invertido(' asd xx}} ')
    self.assertEqual(res, '{{xx')



if __name__ == '__main__':
  suite = unittest.TestLoader().loadTestsFromTestCase(TestXlTex)
  unittest.TextTestRunner(stream=None, verbosity=3).run(suite)