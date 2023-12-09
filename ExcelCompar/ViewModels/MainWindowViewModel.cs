using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocumentFormat.OpenXml.InkML;
using ExcelCompar.Models;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace ExcelCompar.ViewModels
{
    internal class MainWindowViewModel : ObservableObject
    {
        public RelayCommand<string> OpenFileDialogFirstCommand { get; }
        public RelayCommand OpenFileDialogSecondCommand { get; }
        public RelayCommand ReadDataExcelCommand { get; }
        public RelayCommand ComparisonDataCommand { get; }
        LoaderScreen loaderScreen = null;
        public MainWindowViewModel()
        {
            OpenFileDialogFirstCommand = new RelayCommand<string>(OpenFileDialogFirst);
            ReadDataExcelCommand = new RelayCommand(SaveToFile);
            ComparisonDataCommand = new RelayCommand(ReadAlFiles);
        }

        private void OpenFileDialogFirst(string fileNum)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel files (*.xlsx, *.xls)|*.xlsx;*.xls";
            ofd.InitialDirectory = Environment.CurrentDirectory + "\\Temp";
            if ((bool)ofd.ShowDialog())
            {
                if (fileNum == "1")
                    FilePath1 = ofd.FileName;
                if (fileNum == "2")
                    FilePath2 = ofd.FileName;
            }
        }    

        private string _FilePath1;
        public string FilePath1
        {
            get { return _FilePath1; }
            set
            {
                _FilePath1 = value;
                OnPropertyChanged(nameof(FilePath1));
            }
        }

        private string _FilePath2;
        public string FilePath2
        {
            get { return _FilePath2; }
            set
            {
                _FilePath2 = value;
                OnPropertyChanged(nameof(FilePath2));
            }
        }

        private ObservableCollection<Model> _DataFirstList = new ObservableCollection<Model>();
        public ObservableCollection<Model> DataFirstList
        {
            get
            {
                return _DataFirstList;
            }
            set
            {
                _DataFirstList = value;
                OnPropertyChanged(nameof(DataFirstList));
            }
        }

        ////////////////
        private ObservableCollection<Model> _DataSecondtList = new ObservableCollection<Model>();
        public ObservableCollection<Model> DataSecondList
        {
            get
            {
                return _DataSecondtList;
            }
            set
            {
                _DataSecondtList = value;
                OnPropertyChanged(nameof(DataSecondList));
            }
        }
        /////////////////
        private ObservableCollection<Model> _ResultDataList = new ObservableCollection<Model>();
        public ObservableCollection<Model> ResultDataList
        {
            get
            {
                return _ResultDataList;
            }
            set
            {
                _ResultDataList = value;
                OnPropertyChanged(nameof(ResultDataList));
            }
        }


        private void ReadAlFiles()
        {
            loaderScreen = new LoaderScreen();
            loaderScreen.Show();
            loaderScreen.ContentRendered += LoaderScreen_ContentRendered;
        }

        private void LoaderScreen_ContentRendered(object sender, EventArgs e)
        {
            ReadDataFirstFile();
            ReadDataSecondFile();
            СomparisonDataLists();
            loaderScreen.Close();
            loaderScreen = null;
        }

        private void ReadDataFirstFile()
        {
            if (FilePath1 != null & FilePath2 != null)
            {
                DataFirstList.Clear();
                using (XLWorkbook wb = new XLWorkbook(FilePath1))
                {
                    var rows = wb.Worksheet(1).RangeUsed().RowsUsed();
                    foreach (var item in rows)
                    {
                        Model model = new Model()
                        {
                            FirstName = item.Cell(2).Value.ToString(),
                            SecondName = item.Cell(3).Value.ToString(),
                            ThirdName = item.Cell(4).Value.ToString(),
                            Birth = item.Cell(5).Value.ToString().Replace("0:00:00", ""),
                            SMO = item.Cell(6).Value.ToString(),
                            ENP = item.Cell(7).Value.ToString(),
                            CodeMO = item.Cell(8).Value.ToString(),
                            MO = item.Cell(9).Value.ToString(),
                            GroupDN = item.Cell(10).Value.ToString(),
                            NameDN = item.Cell(11).Value.ToString(),
                            City = item.Cell(12).Value.ToString(),
                            Street = item.Cell(13).Value.ToString(),
                            Home = item.Cell(14).Value.ToString(),
                            Corpus = item.Cell(15).Value.ToString(),
                            Room = item.Cell(16).Value.ToString(),
                            RegCity = item.Cell(17).Value.ToString(),
                            RegStreet = item.Cell(18).Value.ToString(),
                            RegHome = item.Cell(19).Value.ToString(),
                            RegCorpus = item.Cell(20).Value.ToString(),
                            RegRoom = item.Cell(21).Value.ToString()
                        };
                        DataFirstList.Add(model);
                    }
                }
            }
        }

        private void ReadDataSecondFile()
        {
            if (FilePath1 != null & FilePath2 != null)
            {
                DataSecondList.Clear();
                using (XLWorkbook wb = new XLWorkbook(FilePath2))
                {
                    var rows = wb.Worksheet(1).RangeUsed().RowsUsed();
                    foreach (var item in rows)
                    {
                        Model model = new Model()
                        {
                            FirstName = item.Cell(2).Value.ToString(),
                            SecondName = item.Cell(3).Value.ToString(),
                            ThirdName = item.Cell(4).Value.ToString(),
                            Birth = item.Cell(5).Value.ToString().Replace("0:00:00", ""),
                            SMO = item.Cell(6).Value.ToString(),
                            ENP = item.Cell(7).Value.ToString(),
                            CodeMO = item.Cell(8).Value.ToString(),
                            MO = item.Cell(9).Value.ToString(),
                            GroupDN = item.Cell(10).Value.ToString(),
                            NameDN = item.Cell(11).Value.ToString(),
                            City = item.Cell(12).Value.ToString(),
                            Street = item.Cell(13).Value.ToString(),
                            Home = item.Cell(14).Value.ToString(),
                            Corpus = item.Cell(15).Value.ToString(),
                            Room = item.Cell(16).Value.ToString(),
                            RegCity = item.Cell(17).Value.ToString(),
                            RegStreet = item.Cell(18).Value.ToString(),
                            RegHome = item.Cell(19).Value.ToString(),
                            RegCorpus = item.Cell(20).Value.ToString(),
                            RegRoom = item.Cell(21).Value.ToString()
                        };
                        DataSecondList.Add(model);
                    }

                }
            }
        }

        private string DateConvert(string input)
        {
            return input.Replace("0:00:00", "");
        }

        private List<Model> _ExcepList = new List<Model>();
        public List<Model> ExcepList
        {
            get
            {
                return _ExcepList;
            }
            set
            {
                _ExcepList = value;
                OnPropertyChanged(nameof(ExcepList));
            }
        }

        private bool _isSave;
        public bool IsSave
        {
            get
            {
                return _isSave;
            }
            set
            {
                _isSave = value;
                OnPropertyChanged(nameof(IsSave));
            }
        }

        private void СomparisonDataLists()
        {
            ExcepList.Clear();
            ExcepList = DataFirstList.Except(DataSecondList, new ModelComparer()).ToList();
            if( ExcepList.Count > 0 )
            {
                IsSave = true;
            }
            else
            {
                IsSave = false;
            }
        }

        private bool _isFormat;
        public bool IsFormat
        {
            get
            {
                return _isFormat;
            }
            set
            {
                _isFormat = value;
                OnPropertyChanged(nameof(IsFormat));
            }
        }

        private void SaveToFile()
        {
            XLWorkbook workbook = new XLWorkbook();
            IXLWorksheet ws = workbook.Worksheets.Add("Data");
            int lenJ = 82;
            if (IsFormat)
            {
                ws.Row(1).Height = 80;
                ws.Row(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                ws.Row(1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                ws.Row(1).Style.Alignment.WrapText = true;
                ws.Column("A").Width = 18;
                ws.Column("B").Width = 15;
                ws.Column("C").Width = 17;
                ws.Column("D").Width = 11;
                ws.Column("E").Width = 8;
                ws.Column("F").Width = 17;
                ws.Column("G").Width = 8;
                ws.Column("H").Width = 14;
                ws.Column("I").Width = 8;
                ws.Column("J").Width = lenJ;
                ws.Column("K").Width = 17;
                ws.Column("L").Width = 18;
                ws.Column("M").Width = 10;
                ws.Column("N").Width = 10;
                ws.Column("O").Width = 10;
                ws.Column("P").Width = 17;
                ws.Column("Q").Width = 18;
                ws.Column("R").Width = 10;
                ws.Column("S").Width = 10;
                ws.Column("T").Width = 10;

                ws.Cell(1, 1).Value = "Фамилия";
                ws.Cell(1, 2).Value = "Имя";
                ws.Cell(1, 3).Value = "Отчество";
                ws.Cell(1, 4).Value = "Дата рождения";
                ws.Cell(1, 5).Value = "СМО";
                ws.Cell(1, 6).Value = "ЕНП";
                ws.Cell(1, 7).Value = "МО прикрепления код";
                ws.Cell(1, 8).Value = "МО прикрепления";
                ws.Cell(1, 9).Value = "Группа диагнозов ДН";
                ws.Cell(1, 10).Value = "Наименование группы диагнозов ДН";
                ws.Cell(1, 11).Value = "Адрес места жительства.Населенный пункт";
                ws.Cell(1, 12).Value = "Адрес места жительства.Улица";
                ws.Cell(1, 13).Value = "Адрес места жительства.Дом";
                ws.Cell(1, 14).Value = "Адрес места жительства.Корпус";
                ws.Cell(1, 15).Value = "Адрес места жительства.Квартира";
                ws.Cell(1, 16).Value = "Адрес места регистрации.Населенный пункт";
                ws.Cell(1, 17).Value = "Адрес места регистрации.Улица";
                ws.Cell(1, 18).Value = "Адрес места регистрации.Дом";
                ws.Cell(1, 19).Value = "Адрес места регистрации.Корпус";
                ws.Cell(1, 20).Value = "Адрес места регистрации.Квартира";
            }

            for (int i = 0; i < ExcepList.Count; i++)
            {
                ws.Cell(i + 2, 1).Value = ExcepList[i].FirstName;
                ws.Cell(i + 2, 2).Value = ExcepList[i].SecondName;
                ws.Cell(i + 2, 3).Value = ExcepList[i].ThirdName;
                ws.Cell(i + 2, 4).Value = ExcepList[i].Birth;
                ws.Cell(i + 2, 5).Value = ExcepList[i].SMO;
                ws.Cell(i + 2, 6).Value = ExcepList[i].ENP;
                ws.Cell(i + 2, 7).Value = ExcepList[i].CodeMO;
                ws.Cell(i + 2, 8).Value = ExcepList[i].MO;
                ws.Cell(i + 2, 9).Value = ExcepList[i].GroupDN;
                ws.Cell(i + 2, 10).Value = ExcepList[i].NameDN;
                ws.Cell(i + 2, 11).Value = ExcepList[i].City;
                ws.Cell(i + 2, 12).Value = ExcepList[i].Street;
                ws.Cell(i + 2, 13).Value = ExcepList[i].Home;
                ws.Cell(i + 2, 14).Value = ExcepList[i].Corpus;
                ws.Cell(i + 2, 15).Value = ExcepList[i].Room;
                ws.Cell(i + 2, 16).Value = ExcepList[i].RegCity;
                ws.Cell(i + 2, 17).Value = ExcepList[i].RegStreet;
                ws.Cell(i + 2, 18).Value = ExcepList[i].RegHome;
                ws.Cell(i + 2, 19).Value = ExcepList[i].RegCorpus;
                ws.Cell(i + 2, 20).Value = ExcepList[i].RegRoom;

                lenJ = ExcepList[i].NameDN.Length;
            }
            ws.Column("J").Width = lenJ + 5;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    workbook.SaveAs(saveFileDialog.FileName);
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error",MessageBoxButton.OK,MessageBoxImage.Error);
                }
            }
        }
    }
}
