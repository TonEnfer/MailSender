using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MailSender
{
    class MailVariable
    {
        public struct Variable
        {
            public int number;
            public string organization;
            public string fullName;
            public List<string> email;
        }

        private List<Variable> mailVar = new List<Variable>();

        public Variable GetVariable(int index)
        {
            return mailVar[index];
        }
        public List<Variable> GetVariables()
        {
            return mailVar;
        }

        public void AddVariable(Variable variable)
        {
            mailVar.Add(variable);
        }
        public void AddVariables(List<Variable> variables)
        {
            mailVar = variables;
        }

        public class Reader
        {
            Excel.Application excelApp;
            Excel.Workbook workbook;
            public Reader(string path)
            {
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                try
                {

                    workbook = excelApp.Workbooks.Open(Path.GetFullPath(path));
                }
                catch
                {
                    excelApp.Quit();
                    Console.WriteLine("Не удалось открыть файл {0}", path);
                    throw new System.Exception();
                }
            }
            public List<Variable> ReadVariable()
            {
                List<Variable> v = new List<Variable>();

                if (workbook.Sheets.Count != 0)
                {
                    int totalRow = excelApp.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    for (int row = 2; row < totalRow + 1; row++)
                    {
                        Variable vvv;
                        vvv.email = new List<string>();
                        vvv.number = (int)excelApp.Cells[RowIndex: row, ColumnIndex: 1].Value();
                        vvv.organization = (string)excelApp.Cells[RowIndex: row, ColumnIndex: 2].Value();
                        vvv.fullName = (string)excelApp.Cells[RowIndex: row, ColumnIndex: 3].Value();
                        string tmp = (string)excelApp.Cells[RowIndex: row, ColumnIndex: 4].Value();
                        foreach (string email in tmp.Split(';'))
                        {
                            vvv.email.Add(email);
                        }
                        v.Add(vvv);
                    }
                }
                excelApp.Quit();
                return v;
            }
        }
    }

}
