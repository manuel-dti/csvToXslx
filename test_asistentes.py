import unittest
import os
import csv
from openpyxl import Workbook
from alumnosCsvToXslx_alfabeticamenteAbajo import setFila, switch_column, paint_row

class TestAsistentes(unittest.TestCase):

    def setUp(self):
        self.datos_prueba = [
            ["Juan", "Pérez"],
            ["Ana", "Gómez"],
            ["Luis", "Martínez"],
            ["María", "López"]
        ]

    def test_setFila_contador_0(self):
        fila = 1        
        fila = setFila(0, fila)
        self.assertEqual(fila, 3)

    def test_setFila_contador_2(self):
        fila = 1        
        fila = setFila(2, fila)
        self.assertEqual(fila, 2)

    def test_setFila_contador_8(self):
        fila = 1        
        fila = setFila(8, fila)
        self.assertEqual(fila, 3) 

    def test_switch_column_no_switch(self):
        expected = {'contador_personas':1, 'columna':1, 'fila':1}
        res = switch_column(contador_personas=1, n_half_alu=5, fila=1, columna=1)
        self.assertEqual(res, expected )

    def test_switch_column_switch(self):
        expected = {'contador_personas':0, 'columna':4, 'fila':1}
        res = switch_column(contador_personas=8, n_half_alu=5, fila=3, columna=1)
        self.assertEqual(res, expected )

    def test_paint_row(self):
        workbook_exp = Workbook()
        expected = workbook_exp.active
        fila = 1
        columna = 1
        expected.cell(row=fila, column=columna).value = self.datos_prueba[0][0] # Nombre
        expected.cell(row=fila, column=columna+1).value = self.datos_prueba[0][1] # Apellido 

        workbook_res = Workbook()
        result = workbook_res.active
        result = paint_row(result, self.datos_prueba, fila, columna)

        self.assertEqual(workbook_exp, workbook_res)

if __name__ == "__main__":
    unittest.main()