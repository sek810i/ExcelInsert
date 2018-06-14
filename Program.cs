using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using OfficeOpenXml;
//using OfficeOpenXml.Drawing;
using System.IO;
using MoreLinq;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;


namespace ExcelInsert
{
    internal class Program
    {
        private class table
        {
            public string FILENAME;
            public string d1;
            public string REGISTR_N;
            public string DATA_REGIS;
            public string TYPE_DOCUMENT;
            public string ISSUER;
            public string FIO_PODPISANT;
            public string DOLG_PODPISANT;
            public string STATUS;
            public string OPER;
            public string d2;
            public string SUBJ_NAME;
            public string SUBJ_INN;
            public string SUBJ_OGRN;
            public string SUBJ_SNILS;
            public string d3;
            public string OBJ_CADNUM;
            public string OBJ_USLNUM;
            public string OBJ_ADDR_FIAS;
            public string OBJ_ADDR_TEXT;
            public string OBJ_ADDR_BUILD;
            public string OBJ_FLAT;
            public string OBJ_UNOM;
            public string OBJ_UNKV;
            public string OBJ_NPP;
            public string OBJ_TYPE;
            public string d4;
            public int ID_DOC;
            public string ID_DOC_ROD;
            public string NUM_OSN;
            public string DATE_OSN;
            public string FIO_P;
            public string RESID;
            public string S;
            public string S_ZHIL;
            public string SOSTAVFAMILY;
            public string ID_DOC_OSN;
            public string S_DOP;
            public string STATUS_DOC;
            public string GOROD;
            public string GOROD_RAYON;
            public string STREETNAME;
            public string NUMDOM;
            public string NUMKORP;
            public string NUMSTR;
            public string NUMVLAD;
            public string NUMFLAT;
            public string NUMROOM;
            public string JPG;
            public string TXT;
            public string PDF;
            public string REC_ID;
            public int mainID;
        }

        private class adress
        {
            public int id;
        }



        private static void Main(string[] args)
        {


            var dbstring = new SqlConnectionStringBuilder
            {
                DataSource = "mysql-techno",
                InitialCatalog = "Test_OB",
                IntegratedSecurity = true,
                ConnectTimeout = 999999999
            };

            var sqlconn = new SqlConnection(dbstring.ConnectionString);
            Console.WriteLine("Начало работы");
            //sqlconn.Open();
            var ladress = sqlconn.Query<int>("select id from disMFC order by id").ToList();
            //sqlconn.Close();
            Console.WriteLine("Загрузил адреса");

            foreach (var bigBatch in ladress.Batch(10000))
            {
                string lastaddres = "";
                string firstaddres = "";
                using (ExcelPackage eP = new ExcelPackage())
                {
                    eP.Workbook.Properties.Author = "Elar";
                    eP.Workbook.Properties.Title = "лист";
                    eP.Workbook.Properties.Company = "Elar";
                    //var st = eP.Workbook.Worksheets.Add("Сводная таблица");
                    var sheet = eP.Workbook.Worksheets.Add("лист");
                    sheet.Cells[1, 1].Value = "mainID";
                    sheet.Cells[1, 2].Value = "OBJ_ADDR_TEXT";
                    sheet.Cells[1, 3].Value = "FILENAME";
                    sheet.Cells[1, 4].Value = "1";
                    sheet.Cells[1, 5].Value = "REGISTR_N";
                    sheet.Cells[1, 6].Value = "DATA_REGIS";
                    sheet.Cells[1, 7].Value = "TYPE_DOCUMENT";
                    sheet.Cells[1, 8].Value = "ISSUER";
                    sheet.Cells[1, 9].Value = "FIO_PODPISANT";
                    sheet.Cells[1, 10].Value = "DOLG_PODPISANT";
                    sheet.Cells[1, 11].Value = "STATUS";
                    sheet.Cells[1, 12].Value = "OPER";
                    sheet.Cells[1, 13].Value = "2";
                    sheet.Cells[1, 14].Value = "SUBJ_NAME";
                    sheet.Cells[1, 15].Value = "SUBJ_INN";
                    sheet.Cells[1, 16].Value = "SUBJ_OGRN";
                    sheet.Cells[1, 17].Value = "SUBJ_SNILS";
                    sheet.Cells[1, 18].Value = "3";
                    sheet.Cells[1, 19].Value = "OBJ_CADNUM";
                    sheet.Cells[1, 20].Value = "OBJ_USLNUM";
                    sheet.Cells[1, 21].Value = "OBJ_ADDR_FIAS";
                    sheet.Cells[1, 22].Value = "OBJ_ADDR_BUILD";
                    sheet.Cells[1, 23].Value = "OBJ_FLAT";
                    sheet.Cells[1, 24].Value = "OBJ_UNOM";
                    sheet.Cells[1, 25].Value = "OBJ_UNKV";
                    sheet.Cells[1, 26].Value = "OBJ_NPP";
                    sheet.Cells[1, 27].Value = "OBJ_TYPE";
                    sheet.Cells[1, 28].Value = "4";
                    sheet.Cells[1, 29].Value = "ID_DOC";
                    sheet.Cells[1, 30].Value = "ID_DOC_ROD";
                    sheet.Cells[1, 31].Value = "NUM_OSN";
                    sheet.Cells[1, 32].Value = "DATE_OSN";
                    sheet.Cells[1, 33].Value = "FIO_P";
                    sheet.Cells[1, 34].Value = "RESID";
                    sheet.Cells[1, 35].Value = "S";
                    sheet.Cells[1, 36].Value = "S_ZHIL";
                    sheet.Cells[1, 37].Value = "SOSTAVFAMILY";
                    sheet.Cells[1, 38].Value = "ID_DOC_OSN";
                    sheet.Cells[1, 39].Value = "S_DOP";
                    sheet.Cells[1, 40].Value = "STATUS_DOC";
                    sheet.Cells[1, 41].Value = "GOROD";
                    sheet.Cells[1, 42].Value = "GOROD_RAYON";
                    sheet.Cells[1, 43].Value = "STREETNAME";
                    sheet.Cells[1, 44].Value = "NUMDOM";
                    sheet.Cells[1, 45].Value = "NUMKORP";
                    sheet.Cells[1, 46].Value = "NUMSTR";
                    sheet.Cells[1, 47].Value = "NUMVLAD";
                    sheet.Cells[1, 48].Value = "NUMFLAT";
                    sheet.Cells[1, 49].Value = "NUMROOM";
                    sheet.Cells[1, 50].Value = "JPG";
                    sheet.Cells[1, 51].Value = "TXT";
                    sheet.Cells[1, 52].Value = "PDF";
                    sheet.Cells[1, 53].Value = "REC_ID";

                    int row = 2;

                    // foreach (var rec in batch)
                    // {
                    //sqlconn.Open();
                    var strBatch = String.Join(",", bigBatch.ToArray());

                    var ltable =
                        sqlconn.Query<table>("select * from MFC where mainid in (" + strBatch + ")", commandTimeout: 0)
                            .ToList();
                    ltable.Sort((emp1, emp2) => emp1.mainID.CompareTo(emp2.mainID));
                    firstaddres = ltable.First().OBJ_ADDR_TEXT;


                    lastaddres = ltable.Last().OBJ_ADDR_TEXT;

                    //sqlconn.Close();


                    foreach (var rows in ltable)
                    {
                        //Console.WriteLine("Начинаю писать в Excel ID = {0}", rec);
                        sheet.Cells[row, 1].Value = rows.mainID;
                        sheet.Cells[row, 2].Value = rows.OBJ_ADDR_TEXT;
                        sheet.Cells[row, 3].Value = rows.FILENAME;
                        sheet.Cells[row, 4].Value = rows.d1;
                        sheet.Cells[row, 5].Value = rows.REGISTR_N;
                        sheet.Cells[row, 6].Style.Numberformat.Format = "dd.MM.yyyy";
                        sheet.Cells[row, 6].Value = rows.DATA_REGIS; //Дата
                        sheet.Cells[row, 7].Value = rows.TYPE_DOCUMENT;
                        sheet.Cells[row, 8].Value = rows.ISSUER;
                        sheet.Cells[row, 9].Value = rows.FIO_PODPISANT;
                        sheet.Cells[row, 10].Value = rows.DOLG_PODPISANT;
                        sheet.Cells[row, 11].Value = rows.STATUS;
                        sheet.Cells[row, 12].Value = rows.OPER;
                        sheet.Cells[row, 13].Value = rows.d2;
                        sheet.Cells[row, 14].Value = rows.SUBJ_NAME;
                        sheet.Cells[row, 15].Value = rows.SUBJ_INN;
                        sheet.Cells[row, 16].Value = rows.SUBJ_OGRN;
                        sheet.Cells[row, 17].Value = rows.SUBJ_SNILS;
                        sheet.Cells[row, 18].Value = rows.d3;
                        sheet.Cells[row, 19].Value = rows.OBJ_CADNUM;
                        sheet.Cells[row, 20].Value = rows.OBJ_USLNUM;
                        sheet.Cells[row, 21].Value = rows.OBJ_ADDR_FIAS;
                        sheet.Cells[row, 22].Value = rows.OBJ_ADDR_BUILD;
                        sheet.Cells[row, 23].Value = rows.OBJ_FLAT;
                        sheet.Cells[row, 24].Value = rows.OBJ_UNOM;
                        sheet.Cells[row, 25].Value = rows.OBJ_UNKV;
                        sheet.Cells[row, 26].Value = rows.OBJ_NPP;
                        sheet.Cells[row, 27].Value = rows.OBJ_TYPE;
                        sheet.Cells[row, 28].Value = rows.d4;
                        sheet.Cells[row, 29].Value = rows.ID_DOC;
                        sheet.Cells[row, 30].Value = rows.ID_DOC_ROD;
                        sheet.Cells[row, 31].Value = rows.NUM_OSN;
                        sheet.Cells[row, 32].Style.Numberformat.Format = "dd.MM.yyyy";
                        sheet.Cells[row, 32].Value = rows.DATE_OSN; //Date
                        sheet.Cells[row, 33].Value = rows.FIO_P;
                        sheet.Cells[row, 34].Value = rows.RESID;
                        sheet.Cells[row, 35].Value = rows.S;
                        sheet.Cells[row, 36].Value = rows.S_ZHIL;
                        sheet.Cells[row, 37].Value = rows.SOSTAVFAMILY;
                        sheet.Cells[row, 38].Value = rows.ID_DOC_OSN;
                        sheet.Cells[row, 39].Value = rows.S_DOP;
                        sheet.Cells[row, 40].Value = rows.STATUS_DOC;
                        sheet.Cells[row, 41].Value = rows.GOROD;
                        sheet.Cells[row, 42].Value = rows.GOROD_RAYON;
                        sheet.Cells[row, 43].Value = rows.STREETNAME;
                        sheet.Cells[row, 44].Value = rows.NUMDOM;
                        sheet.Cells[row, 45].Value = rows.NUMKORP;
                        sheet.Cells[row, 46].Value = rows.NUMSTR;
                        sheet.Cells[row, 47].Value = rows.NUMVLAD;
                        sheet.Cells[row, 48].Value = rows.NUMFLAT;
                        sheet.Cells[row, 49].Value = rows.NUMROOM;
                        sheet.Cells[row, 50].Value = rows.JPG;
                        sheet.Cells[row, 51].Value = rows.TXT;
                        sheet.Cells[row, 52].Value = rows.PDF;
                        sheet.Cells[row, 53].Value = rows.REC_ID;
                        if (row%100 == 0)
                        {
                            Console.WriteLine("Записал {0} строк", row);
                        }

                        row++;

                    }

                    //}

                    // var pt = st.PivotTables.Add(st.Cells["A5"],sheet.Cells[1,1,row,53],"qwe");

                    //pt.Compact = true;
                    //pt.CompactData = true;
                    //pt.GrandTotalCaption = "Total amount";
                    //pt.RowFields.Add(pt.Fields[0]);
                    //pt.RowFields.Add(pt.Fields[1]);
                    //pt.PageFields.Add(pt.Fields[1]);
                    //pt.RowFields.Add(pt.Fields[2]);
                    //pt.RowFields.Add(pt.Fields[3]);
                    //pt.RowFields.Add(pt.Fields[4]);
                    //pt.RowFields.Add(pt.Fields[5]);
                    //pt.RowFields.Add(pt.Fields[6]);
                    //pt.RowFields.Add(pt.Fields[7]);
                    //pt.RowFields.Add(pt.Fields[8]);





                    // pt.PageFields.Add(pt.Fields[2]);


                    //pt.DataFields.Add(pt.Fields[2]);
                    //pt.DataFields.Add(pt.Fields[3]);
                    //pt.DataFields.Add(pt.Fields[4]);
                    //pt.DataFields.Add(pt.Fields[5]);
                    //pt.DataFields.Add(pt.Fields[6]);
                    //pt.DataFields.Add(pt.Fields[7]);
                    //pt.DataFields[0].Function = DataFieldFunctions.Product;
                    //pt.DataOnRows = false;

                    var rage = sheet.Cells[1, 1, row, 53];
                    var table = sheet.Tables.Add(rage, "table1");
                    table.ShowTotal = true;
                    table.TableStyle = TableStyles.Light9;
                    Console.WriteLine("Сохраняю Excel");
                    var bin = eP.GetAsByteArray();
                    File.WriteAllBytes(
                        @"d:\!!!!!!!!!!!!!!!!!!!!!!!МФЦ\Excel\" + firstaddres.Replace(@"\", @"_").Replace(@"/", @"_") +
                        " - " + lastaddres.Replace(@"\", @"_").Replace(@"/", @"_") + ".xlsx", bin);
                }
            }
        }
    }
}


