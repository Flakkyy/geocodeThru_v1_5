using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Windows.Threading;
using System.Threading;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using DotSpatial.Projections;

// ReSharper disable PossibleInvalidOperationException
// ReSharper disable RedundantBoolCompare

namespace GeocodeThru
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {

        // ReSharper disable once RedundantDefaultMemberInitializer
        bool _quitClicked = false;
        // ReSharper disable once RedundantDefaultMemberInitializer
        bool _complete = false;
        // counting rows and columns
        int _rowCount;
        int _columnCount;
        // excel path
        string _excelPath;
        Excel.Application _excelApp;
        Excel.Workbook _excelWorkBook;
        Excel._Worksheet _excelWorkSheet;


        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// loading values from header of excel into comboboxes
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menu_open_xlsx_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var excelFile = new OpenFileDialog
                {
                    DefaultExt = ".xlsx",
                    Filter = "(.xlsx)|*.xlsx"
                };
                Nullable<bool> result = excelFile.ShowDialog();
                if (result != true) return;

                RtbProgress.AppendText("Loading Excell file and checking if APIs services are available.\r...\r");
                DoEvents();
                CbxUlice.Items.Clear();
                CbxCp.Items.Clear();
                CbxCo.Items.Clear();
                CbxObec.Items.Clear();
                CbxPsc.Items.Clear();

                _excelPath = excelFile.FileName;

                _excelApp = new Excel.Application();
                _excelWorkBook = _excelApp.Workbooks.Open(_excelPath);
                _excelWorkSheet = _excelWorkBook.Sheets[1];
                // counting rows and columns
                _rowCount = _excelWorkSheet.UsedRange.Rows.Count;
                _columnCount = _excelWorkSheet.UsedRange.Columns.Count;
                // add max rows into interval
                TxtInterval.Text = "[1 - " + _rowCount + "]";
                // Adding selections to comboboxes
                for (int i = 1; i < _excelWorkSheet.UsedRange.Columns.Count + 1; i++)
                {
                    CbxUlice.Items.Add(_excelWorkSheet.Cells[1, i].Value);
                    CbxCp.Items.Add(_excelWorkSheet.Cells[1, i].Value);
                    CbxCo.Items.Add(_excelWorkSheet.Cells[1, i].Value);
                    CbxObec.Items.Add(_excelWorkSheet.Cells[1, i].Value);
                    CbxPsc.Items.Add(_excelWorkSheet.Cells[1, i].Value);
                }
                // adding one null slot to comboboxes
                CbxUlice.Items.Add("");
                CbxCp.Items.Add("");
                CbxCo.Items.Add("");
                CbxObec.Items.Add("");
                CbxPsc.Items.Add("");

                EnableObjects(true);
                MenuOpen.IsEnabled = false;

                RtbProgress.AppendText(_excelPath + " is loaded!\r");
                RtbProgress.AppendText("For new excel file RESTART THE APP please.\r");
                RtbProgress.ScrollToEnd();
                RtbProgress.AppendText("---------------\r");
                CheckIfApisAreOnline();
                RtbProgress.AppendText("---------------\r");
            }
            catch (Exception ex)
            {
                MessageBox.Show("A handled exception just occurred: " + ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        /// <summary>
        /// enabling or disabling all comboboxes
        /// </summary>
        /// <param name="enab"></param>
        private void EnableObjects(bool enab)
        {
            MenuOpenXlsx.IsEnabled = enab;
            //MenuOpenAccdb.IsEnabled = enab;
            MItemStart.IsEnabled = enab;

            GrbApis.IsEnabled = enab;
            GrbKeys.IsEnabled = enab;
            GrbColMap.IsEnabled = enab;

            TxtbxFromRow.IsEnabled = enab;


        }

        public void DoEvents()
        {
            try
            {
                DispatcherFrame frame = new DispatcherFrame();
                Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background,
                    new DispatcherOperationCallback(ExitFrames), frame);
                Dispatcher.PushFrame(frame);
            }
            catch (Exception ex)
            {
                MessageBox.Show("A handled exception just occurred: " + ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        public object ExitFrames(object f)
        {
            ((DispatcherFrame)f).Continue = false;

            return null;
        }
        /// <summary>
        /// do all other work
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mItem_Start_Click(object sender, RoutedEventArgs e)
        {
            Task.Factory.StartNew(StartGeocoding);
            //Task.Factory.StartNew(StartGeocoding);
            //Thread startBtnGeocod = new Thread(StartGeocoding);
            //startBtnGeocod.Start();
        }

        public void StartGeocoding()
        {
            try
            {
                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                //delegate
                {
                    if (CbxRuain.IsChecked.Value == true || CbxOsm.IsChecked.Value == true || CbxMq.IsChecked.Value == true || CbxMcz.IsChecked.Value == true || CbxHm.IsChecked.Value == true || CbxGm.IsChecked.Value == true || CbxBm.IsChecked.Value == true)
                    {
                        if (CbxUlice.Text.Equals("") && CbxCp.Text.Equals("") && CbxCo.Text.Equals("") && CbxObec.Text.Equals("")
                            && CbxPsc.Text.Equals(""))
                        {
                            WarningMsg("Fill atleast one drop-down menu!");
                            FlashRectangle(RectDropdowns);
                        }
                        else
                        {
                            try
                            {
                                RtbProgress.AppendText("[" + DateTime.Now.ToLongTimeString() + "] STARTING GEOCODING\r");
                                DoEvents();
                                EnableObjects(false);
                                _excelWorkBook.Unprotect();
                                _excelWorkSheet.Unprotect();
                                // getting indexes of columns for each selected columns
                                int[] indexy = new int[5];
                                for (int c = 1; c <= _columnCount; c++)
                                {
                                    if (_excelWorkSheet.Cells[1, c].Value.ToString() == CbxUlice.Text) { indexy[0] = c; }
                                    if (_excelWorkSheet.Cells[1, c].Value.ToString() == CbxCp.Text) { indexy[1] = c; }
                                    if (_excelWorkSheet.Cells[1, c].Value.ToString() == CbxCo.Text) { indexy[2] = c; }
                                    if (_excelWorkSheet.Cells[1, c].Value.ToString() == CbxObec.Text) { indexy[3] = c; }
                                    if (_excelWorkSheet.Cells[1, c].Value.ToString() == CbxPsc.Text) { indexy[4] = c; }
                                }
                                // write a header for coordinates, cbx 
                                int i = 0;
                                foreach (var item in GetBoxes())
                                {
                                    CheckBox cbx = item as CheckBox;
                                    // ReSharper disable once PossibleNullReferenceException
                                    if (cbx.Content.ToString() == "RUIAN")
                                    {
                                        _excelWorkSheet.Cells[1, _columnCount + i + 1] = "X_" + cbx.Content;
                                        _excelWorkSheet.Cells[1, _columnCount + i + 2] = "Y_" + cbx.Content;
                                        i += 2;
                                    }
                                    else
                                    {
                                        _excelWorkSheet.Cells[1, _columnCount + i + 1] = "lat_" + cbx.Content;
                                        _excelWorkSheet.Cells[1, _columnCount + i + 2] = "lon_" + cbx.Content;
                                        _excelWorkSheet.Cells[1, _columnCount + i + 3] = "X_" + cbx.Content;
                                        _excelWorkSheet.Cells[1, _columnCount + i + 4] = "Y_" + cbx.Content;
                                        i += 4;
                                    }

                                }
                                // this sets starting row 
                                int startingRow = Convert.ToInt32(TxtbxFromRow.Text);
                                if (startingRow == 1) startingRow = 2;
                                // this list store coordiantes for further writing to excel sheet
                                List<double> coords = new List<double>();

                                #region urls and keys for APIs
                                string urlRuian = "http://www.vugtk.cz/euradin/services/rest.py/Geocode/text?SearchText={0}&SuppressID=off";
                                string keyGm = TxtboxGmKey.Text;
                                // console.developers.google.com/project/sixth-shield-109215/apiui/credential
                                string urlGoogle = "https://maps.googleapis.com/maps/api/geocode/xml?address={0}&key=" + keyGm;

                                string hmAppId = TxtboxHmAppId.Text; //"O5Xux7fAgmj4kSi67XbA";
                                string hmAppCode = TxtboxHmAppCode.Text; //"cbeXpCRSVWo9kc17HbtHEA";
                                string urlHereMaps = "http://geocoder.cit.api.here.com/6.2/geocode.xml?app_id=" + hmAppId + "&app_code=" + hmAppCode + "&gen=9&searchtext={0}";

                                string keyMapQ = TxtboxMqKey.Text; //"2i74YRgMWpE5GJoOlkFpy57yINNjQQ1V";
                                string urlMapQ = "http://open.mapquestapi.com/geocoding/v1/address?key=" + keyMapQ + "&outFormat=xml&maxResults=1&thumbMaps=false&location={0}";

                                string urlMcz = "http://api.mapy.cz/geocode?query={0}";
                                string urlOsm = "http://nominatim.openstreetmap.org/search/cz/{0}?format=xml&polygon=0&addressdetails=1&limit=1";

                                string keyBingM = TxtboxBmKey.Text; //"wrIA0ucuQwsQUxP6OAZP~GcpFgalVzrLfG6E-qgBnaQ~AhcLmtVfz7TlpdyF12sCtimnxcho0RXl_eW_FRJIYlDjaiGiq-a1lc2cZOxmKBIb";
                                                                    // msdn.microsoft.com/en-us/library/gg650598.aspx
                                string urlBm = "http://dev.virtualearth.net/REST/v1/Locations?q={0}&o=xml&maxRes=1&key=" + keyBingM;
                                #endregion

                                for (int r = startingRow; r <= _rowCount; r++)
                                {
                                    // if quit is clicked close the app
                                    if (_quitClicked == true)
                                    {
                                        MItemQuit.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
                                    }
                                    // printing which record is processing now
                                    RtbProgress.AppendText("-----------------------------------------------------------------\r");
                                    RtbProgress.AppendText(
                                        "[" + DateTime.Now.ToLongTimeString() + $"] Processing record {r.ToString()} out of {_rowCount.ToString()}\r");
                                    RtbProgress.ScrollToEnd();
                                    DoEvents();
                                    // getting values of each cells
                                    #region values of cells
                                    string ulice;
                                    string cp;
                                    string co;
                                    string obec;
                                    string psc;

                                    if (indexy[0] == 0) ulice = "";
                                    else
                                    {
                                        if (_excelWorkSheet.Cells[r, indexy[0]].Value == null) ulice = "";
                                        else
                                        {
                                            ulice = _excelWorkSheet.Cells[r, indexy[0]].Value.ToString().Trim();
                                            ulice = TrimLastChar(ulice);
                                            ulice = ReplaceCharacters(ulice, "nám.", "náměstí ");
                                            ulice = RemoveStrings(ulice);
                                            ulice = RemoveAllAfterFirst(ulice, "-");
                                            ulice = RemoveAllAfterFirst(ulice, "/");

                                        }
                                    }
                                    if (indexy[1] == 0) cp = "";
                                    else
                                    {
                                        if (_excelWorkSheet.Cells[r, indexy[1]].Value == null) cp = "";
                                        else
                                        {
                                            cp = _excelWorkSheet.Cells[r, indexy[1]].Value.ToString().Trim();
                                            cp = RemoveAllAfterFirst(cp, "-");
                                            cp = RemoveAllAfterFirst(cp, "/");
                                            cp = TrimLastChar(cp);
                                        }
                                    }
                                    if (indexy[2] == 0) co = "";
                                    else
                                    {
                                        if (_excelWorkSheet.Cells[r, indexy[2]].Value == null) co = "";
                                        else
                                        {
                                            co = _excelWorkSheet.Cells[r, indexy[2]].Value.ToString().Trim();
                                            co = RemoveAllAfterFirst(co, "-");
                                            co = RemoveAllAfterFirst(co, "/");
                                            co = TrimLastChar(co);
                                        }
                                    }
                                    if (indexy[3] == 0) obec = "";
                                    else
                                    {
                                        if (_excelWorkSheet.Cells[r, indexy[3]].Value == null) obec = "";
                                        else
                                        {
                                            obec = _excelWorkSheet.Cells[r, indexy[3]].Value.ToString().Trim();
                                            obec = RemoveDigits(obec);
                                            obec = RemoveAllAfterFirst(obec, "-");
                                            obec = RemoveStrings(obec);
                                        }
                                    }
                                    if (indexy[4] == 0) psc = "";
                                    else
                                    {
                                        if (_excelWorkSheet.Cells[r, indexy[4]].Value == null) psc = "";
                                        else
                                        {
                                            psc = _excelWorkSheet.Cells[r, indexy[4]].Value.ToString().Trim();
                                            psc = ReplaceCharacters(psc, " ", "");
                                        }
                                    }
                                    #endregion
                                    RtbProgress.AppendText("Adress: " + (ulice + " " + cp + "/" + co + " " + obec + " " + psc).Trim() + "\r");
                                    RtbProgress.ScrollToEnd();
                                    #region ruain geocoding
                                    if (CbxRuain.IsChecked.Value == true)
                                    {
                                        string geocodeRowRuian = (ulice + " " + cp + " " + co + " " + obec + " " + psc).Trim();
                                        if (geocodeRowRuian == "") continue;
                                        string[] xy;
                                        // calling webclient for geocoding thru ruian
                                        string s = DownloadCoords(urlRuian, geocodeRowRuian);
                                        // determining that if s has value

                                        if (s == "")
                                        {
                                            geocodeRowRuian = (ulice + " " + co + " " + obec + " " + psc).Trim();
                                            s = DownloadCoords(urlRuian, geocodeRowRuian);
                                            if (s == "")
                                            {
                                                geocodeRowRuian = (ulice + " " + cp + " " + obec + " " + psc).Trim();
                                                s = DownloadCoords(urlRuian, geocodeRowRuian);
                                                if (s == "")
                                                {
                                                    geocodeRowRuian = (ulice + " " + obec + " " + psc).Trim();
                                                    s = DownloadCoords(urlRuian, geocodeRowRuian);
                                                    xy = SplitSourceString(s);
                                                }
                                                else { xy = SplitSourceString(s); }
                                            }
                                            else { xy = SplitSourceString(s); }
                                        }
                                        else { xy = SplitSourceString(s); }

                                        if (s == "")
                                        {
                                            coords.Add(0);
                                            coords.Add(0);
                                        }
                                        else
                                        {
                                            // SJTSK coords
                                            coords.Add(Convert.ToDouble(xy[0].Insert(0, "-").Replace(".", ",")));
                                            coords.Add(Convert.ToDouble(xy[1].Trim().Insert(0, "-").Replace(".", ",")));
                                        }
                                    }
                                    #endregion
                                    #region GoogleMaps
                                    // developers.google.com/maps/documentation/geocoding/intro
                                    if (CbxGm.IsChecked.Value == true)
                                    {
                                        string geocodeGm = (ulice + " " + cp + " " + co + " " + obec + " " + psc).Trim();
                                        if (geocodeGm == "") continue;
                                        string s = DownloadCoords(urlGoogle, geocodeGm);
                                        if (RegexBetween(s, "<status>(.*)</status>").Equals("OK"))
                                        {
                                            s = RemoveAllAfterFirst(s, "</location>");
                                            CreateCoordsAndTransform(s, coords, "<lng>(.*)</lng>", "<lat>(.*)</lat>");
                                        }
                                        else { AddNulls(coords); }
                                    }
                                    #endregion
                                    #region HereMaps
                                    // developer.here.com/rest-apis/documentation/geocoder/topics/request-constructing.html
                                    //developer.here.com/documentation/download/geocoding_nlp/6.2.91/Geocoder%20API%20v6.2.91%20Developer's%20Guide.pdf
                                    if (CbxHm.IsChecked.Value == true)
                                    {
                                        string geocodeHm = (ulice + " " + cp + " " + co + " " + obec + " " + psc).Trim();
                                        if (geocodeHm == "") continue;
                                        string s = DownloadCoords(urlHereMaps, geocodeHm);
                                        if (RegexBetween(s, "<ViewId>(.*)</ViewId>").Equals("0"))
                                        {
                                            string partS = RegexBetween(s, "<DisplayPosition>(.*)</DisplayPosition>");
                                            CreateCoordsAndTransform(partS, coords, "<Longitude>(.*)</Longitude>", "<Latitude>(.*)</Latitude>");
                                        }
                                        else { AddNulls(coords); }
                                    }
                                    #endregion
                                    #region MapQuest
                                    // open.mapquestapi.com/geocoding
                                    if (CbxMq.IsChecked.Value == true)
                                    {
                                        string geocodeMq = (ulice + " " + cp + " " + co + " " + obec + " " + psc).Trim();
                                        if (geocodeMq == "") continue;
                                        string s = DownloadCoords(urlMapQ, geocodeMq);
                                        if (RegexBetween(s, "<statusCode>(.*)</statusCode>").Equals("0"))
                                        {
                                            s = RemoveAllAfterFirst(s, "</latLng>");
                                            CreateCoordsAndTransform(s, coords, "<lng>(.*)</lng>", "<lat>(.*)</lat>");
                                        }
                                        else { AddNulls(coords); }
                                    }
                                    #endregion
                                    #region MapyCZ
                                    // api.mapy.cz
                                    if (CbxMcz.IsChecked.Value == true)
                                    {
                                        string geocodeMcz = (ulice + " " + cp + " " + co + " " + obec + " " + psc).Trim();
                                        if (geocodeMcz == "") continue;
                                        string s = DownloadCoords(urlMcz, geocodeMcz);
                                        if (RegexBetween(s, "message=\"(.*)\" >").Equals("OK") && s.Contains("<item"))
                                        {
                                            CreateCoordsAndTransform(s, coords, "x=\"(.*)\"\n", "y=\"(.*)\"\n");
                                        }
                                        else { AddNulls(coords); }
                                    }
                                    #endregion
                                    #region OSM
                                    // wiki.openstreetmap.org/wiki/Nominatim
                                    if (CbxOsm.IsChecked.Value == true)
                                    {
                                        string geocodeOsm = (ulice + " " + cp + " " + co + " " + obec + " " + psc).Trim();
                                        if (geocodeOsm == "") continue;
                                        string s = DownloadCoords(urlOsm, geocodeOsm);
                                        if (RegexBetween(s, "class='(.*)' type").Equals("place"))
                                        {
                                            s = RemoveAllAfterFirst(s, " display_name=");
                                            CreateCoordsAndTransform(s, coords, "lon='(.*)'", "lat='(.*)' lon=");
                                        }
                                        else { AddNulls(coords); }
                                    }
                                    #endregion
                                    #region BingMaps
                                    // msdn.microsoft.com/en-us/library/ff701711.aspx
                                    if (CbxBm.IsChecked.Value == true)
                                    {
                                        string geocodeBm = (ulice + " " + cp + " " + co + " " + obec + " " + psc).Trim();
                                        if (geocodeBm == "") continue;
                                        string s = DownloadCoords(urlBm, geocodeBm);
                                        if (RegexBetween(s, "<StatusDescription>(.*)</StatusDescription>").Equals("OK") &&
                                            RegexBetween(s, "<EstimatedTotal>(.*)</EstimatedTotal>").Equals("1"))
                                        {
                                            s = RemoveAllAfterFirst(s, "</Point>");
                                            CreateCoordsAndTransform(s, coords, "<Longitude>(.*)</Longitude>", "<Latitude>(.*)</Latitude>");
                                        }
                                        else { AddNulls(coords); }
                                    }
                                    #endregion

                                    #region write values to row for all APIs
                                    for (int j = 0; j < coords.Count; j++)
                                    {
                                        // ReSharper disable once CompareOfFloatsByEqualityOperator
                                        if (coords[j] == 0)
                                        {
                                            _excelWorkSheet.Cells[r, _columnCount + j + 1].Value = " ";
                                        }
                                        else
                                        {
                                            _excelWorkSheet.Cells[r, _columnCount + j + 1].Value = coords[j];
                                        }
                                    }
                                    #endregion
                                    coords.Clear();
                                    // updating progress bar and progress text on groupbox                                        
                                    float hodnota = r / (float)_rowCount;
                                    ProgressBar.Value = hodnota * 100;
                                    GrbProgress.Header = "Progress... " + Math.Round((hodnota * 100), 2) + " %";
                                    DoEvents();

                                    _excelWorkBook.Save();
                                    Thread.Sleep(1100);
                                }
                                // ReSharper disable once RedundantAssignment
                                indexy = null;

                                EnableObjects(true);
                                RtbProgress.AppendText("-----------------\r");
                                RtbProgress.AppendText("| COMPLETED! |\r");
                                RtbProgress.AppendText("-----------------\r");
                                RtbProgress.AppendText("PLEASE RESET THE APPLICATION, IF YOU WANT TO USE ANOTHER GEOCODING.\r");
                                RtbProgress.ScrollToEnd();

                                MItemStart.IsEnabled = false;
                                MenuOpen.IsEnabled = false;
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("A handled exception just occurred: " + ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Warning);
                            }
                            // Close excel
                            //_excelWorkBook.Protect();
                            //_excelWorkSheet.Protect();
                            CloseExcel();
                            _complete = true;
                        }
                    }
                    else
                    {
                        WarningMsg("You didn't chosse any method for geocoding!");
                        FlashRectangle(RectMethods);
                    }
                }
            ));
            }
            catch (Exception ex)
            {
                MessageBox.Show("A handled exception just occurred: " + ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        /// <summary>
        /// take a input coordinates, write them ass WGS and transform them into JSTSK and also write them
        /// </summary>
        /// <param name="strSource">downloaded string from API</param>
        /// <param name="coordsList">list of coordiantes that will be writen to excel file</param>
        /// <param name="regexLot">get a longitude between this string</param>
        /// <param name="regexLat">get a latitude between this string</param>
        private void CreateCoordsAndTransform(string strSource, List<double> coordsList, string regexLot, string regexLat)
        {
            try
            {
                //Defines the starting coordiante system
                ProjectionInfo pStart = KnownCoordinateSystems.Geographic.World.WGS1984;
                //Defines the ending coordiante system
                ProjectionInfo pEnd = KnownCoordinateSystems.Projected.NationalGrids.SJTSKKrovakEastNorth;
                // define z coords for tranformation
                double[] z = new double[] { 1 };
                double[] xy = new double[] { Convert.ToDouble(RegexBetween(strSource,regexLot)),
                                          Convert.ToDouble(RegexBetween(strSource, regexLat))};
                // adding lan and lon to the list 
                coordsList.Add(xy[1]);
                coordsList.Add(xy[0]);
                //Calls the reproject function that will transform the input location to the output locaiton
                Reproject.ReprojectPoints(xy, z, pStart, pEnd, 0, 1);
                // adding transformated (to sjtsk) lat and lot to list
                coordsList.Add(xy[0]);
                coordsList.Add(xy[1]);
            }
            catch (Exception ex)
            {
                MessageBox.Show("A handled exception just occurred: " + ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        /// <summary>
        /// just add 4 zeros to list of coordinates
        /// </summary>
        /// <param name="coords">input list of coordinates</param>
        private void AddNulls(List<double> coords)
        {
            for (int iadd = 0; iadd < 4; iadd++)
            {
                coords.Add(0);
            }
        }
        /// <summary>
        /// get checkboxes as object for furter writing into excel file
        /// </summary>
        /// <returns>list of checkboxes that are active/true</returns>
        List<object> GetBoxes()
        {
            List<object> checkBoxes = new List<object>();
            if (CbxRuain.IsChecked.Value == true) checkBoxes.Add(CbxRuain);
            if (CbxGm.IsChecked.Value == true) checkBoxes.Add(CbxGm);
            if (CbxHm.IsChecked.Value == true) checkBoxes.Add(CbxHm);
            if (CbxMq.IsChecked.Value == true) checkBoxes.Add(CbxMq);
            if (CbxMcz.IsChecked.Value == true) checkBoxes.Add(CbxMcz);
            if (CbxOsm.IsChecked.Value == true) checkBoxes.Add(CbxOsm);
            if (CbxBm.IsChecked.Value == true) checkBoxes.Add(CbxBm);
            return checkBoxes;
        }
        /// <summary>
        /// get coordinate between two strings
        /// </summary>
        /// <param name="inputText">downloaded string from API</param>
        /// <param name="find">get a string between this string</param>
        /// <returns></returns>
        string RegexBetween(string inputText, string find)
        {
            Regex reg = new Regex(find);
            string result = reg.Match(inputText).Groups[1].ToString().Replace(".", ",");
            if (result == "") result = "0";
            return result;
        }
        /// <summary>
        /// flashing area that needs to be filled
        /// </summary>
        /// <param name="rect">input area/rectangle</param>
        private void FlashRectangle(object rect)
        {
            Rectangle recta = rect as Rectangle;
            for (int i = 0; i < 3; i++)
            {
                // ReSharper disable once PossibleNullReferenceException
                recta.Visibility = Visibility.Visible;
                DoEvents();
                Thread.Sleep(300);
                DoEvents();
                recta.Visibility = Visibility.Hidden;
                DoEvents();
                Thread.Sleep(300);
                DoEvents();
            }
        }
        /// <summary>
        /// Warning message
        /// </summary>
        /// <param name="msg">send message</param>
        private void WarningMsg(string msg)
        {
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Warning;
            MessageBox.Show(msg, "Warning", button, icon);
        }
        /// <summary>
        /// simple split of downloaded string from ruian
        /// </summary>
        /// <param name="input">string coordinates</param>
        /// <returns></returns>
        string[] SplitSourceString(string input)
        {
            string[] xy;
            if (CountLinesInString(input) > 1)
            {
                // ReSharper disable once StringIndexOfIsCultureSpecific.1
                input = input.Substring(0, input.IndexOf(Environment.NewLine));
                xy = input.Split(new[] { "," }, StringSplitOptions.None);
            }
            else
            {
                xy = input.Split(new[] { "," }, StringSplitOptions.None);
            }
            return xy;
        }

        /// <summary>
        /// download coordiantes from APIs and return them
        /// </summary>
        /// <param name="url">input url in string</param>
        /// <param name="inputSearch">searched text</param>
        /// <returns></returns>
        string DownloadCoords(string url, string inputSearch)
        {
            WebDownload wc = new WebDownload { Timeout = 15000 };
            string s = wc.DownloadString(String.Format(url, inputSearch));
            return s;
        }
        /// <summary>
        /// remove digits from string
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        string RemoveDigits(string input)
        {
            string result = Regex.Replace(input, @"[\d]", "");
            return result.Trim();
        }
        /// <summary>
        /// remove all characters after first input character
        /// </summary>
        /// <param name="input"></param>
        /// <param name="character"></param>
        /// <returns></returns>
        string RemoveAllAfterFirst(string input, string character)
        {
            int index = input.IndexOf(character, StringComparison.Ordinal);
            if (index > 0)
                input = input.Substring(0, index);
            return input.Trim();
        }
        /// <summary>
        /// remove(TrimEnd) characters like a, b, c, d from input string
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        string TrimLastChar(string input)
        {
            string result = input;
            if (result.Length > 1)
            {
                string last2 = input.Substring(input.Length - 2);
                if (Regex.IsMatch(last2[0].ToString(), @"\d"))
                {
                    char[] characters = { 'a', 'A', 'b', 'B', 'c', 'C', 'd', 'D', 'e', 'E' };
                    result = input.TrimEnd(characters);
                }
            }
            return result.Trim();

        }
        /// <summary>
        /// it's a just simple replace 
        /// </summary>
        /// <param name="input"></param>
        /// <param name="oldStr"></param>
        /// <param name="newStr"></param>
        /// <returns></returns>
        string ReplaceCharacters(string input, string oldStr, string newStr)
        {
            string result = input.Replace(oldStr, newStr);
            return result.Trim();
        }
        /// <summary>
        /// remove strings like č.p., čp., č.o., čo., e.č. etc..
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        string RemoveStrings(string input)
        {
            string[] str = { "č.p.", "č. p.", "čp.", "čp ", "č.o.", "č. o.", "čo.", "e.č.", "e. č.", "č.e.", "č.ev.", "XXX",
                               "I.", /*"I",*/ "II.", "II", "III.", "III", "IV.", "IV",
                               "parc. č.", "parc. Č.", "p.č.", "st.", "parc.č.","č.parcely", "p.p.č.", "st. p. č.", "ul.", "u.",
                               "\"", "+", ",", "(",")", "|", "´", "¨",
                               "ČS PHM a LPG","ČS PHM","ČS LPG", "Agip","ČS JJ Tank","Benzina","Eurobit","Robin Oil","Hunsgas",
                               "E.O.C - SHELL","ČS CNG","Recap Trade","ČS Shell","Shell","ARMEX Oil","PS CNG","Globus","Zerogas",
                               "Lukoil","Pap Oil","ČS OMV", "OMV", "ČEPRO"};
            foreach (string strItem in str)
            {
                // TODO: how to resolve this? big "I" will be replaced everytime and that can be damaging
                input = input.Replace(strItem, "");
            }
            return input.Trim();
        }
        /// <summary>
        /// counting lines in string
        /// </summary>
        /// <param name="s">input string</param>
        /// <returns></returns>
        static long CountLinesInString(string s)
        {
            long count = 1;
            int start = 0;
            while ((start = s.IndexOf('\n', start)) != -1)
            {
                count++;
                start++;
            }
            return count;
        }
        /// <summary>
        /// just litle about app
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mItem_About_Click(object sender, RoutedEventArgs e)
        {
            AboutBox1 about = new AboutBox1();

            about.ShowDialog();

        }
        /// <summary>
        /// releasing object of excel app
        /// </summary>
        /// <param name="obj">com/excel object to be released</param>
        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                //obj = null;
            }
            catch (Exception ex)
            {
                //obj = null;
                MessageBox.Show("Unable to release the Object " + ex);
            }
            finally
            {
                GC.Collect();
            }
        }
        /// <summary>
        /// close the excel app and all stuff around that
        /// </summary>
        private void CloseExcel()
        {
            try
            {
                if (_excelWorkSheet != null)
                {
                    _excelWorkBook.Save();
                    _excelWorkBook.Close(true, null, null);
                    _excelApp.Quit();
                    ReleaseObject(_excelWorkSheet);
                    ReleaseObject(_excelWorkBook);
                    ReleaseObject(_excelApp);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("A handled exception just occurred: " + ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// close the app
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mItem_Quit_Click(object sender, RoutedEventArgs e)
        {

            _quitClicked = true;
            if (_complete != true)
            {
                CloseExcel();
            }

            Close();
        }

        #region handle textbox From beggining row
        /// <summary>
        /// allowing only number in this textbox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtbxFromRow_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!char.IsDigit(e.Text, e.Text.Length - 1))
            {
                e.Handled = true;
            }
        }
        /// <summary>
        /// make sure that range is always OK, interval from 1 to maxRows
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtbxFromRow_TextChanged(object sender, TextChangedEventArgs e)
        {

            if (TxtbxFromRow.IsEnabled == true)
            {
                if (TxtbxFromRow.Text == "") TxtbxFromRow.Text = "1";
                else if (Convert.ToInt32(TxtbxFromRow.Text) < 1) TxtbxFromRow.Text = "1";
                else if (Convert.ToInt32(TxtbxFromRow.Text) > _rowCount) TxtbxFromRow.Text = _rowCount.ToString();

            }
        }
        /// <summary>
        /// select all text in box when is box get focus
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtbxFromRow_GotFocus(object sender, RoutedEventArgs e)
        {
            ((TextBox)sender).SelectAll();
        }
        /// <summary>
        /// when is mouse released, then select all text in box 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtbxFromRow_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var textBox = (TextBox)sender;
            if (!textBox.IsKeyboardFocusWithin)
            {
                textBox.Focus();
                e.Handled = true;
            }
        }
        #endregion

        /// <summary>
        /// RESET APP IF NEEDED
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void mItem_Reset_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(Application.ResourceAssembly.Location);
                Application.Current.Shutdown();
            }
            catch (Exception ex)
            {
                MessageBox.Show("A handled exception just occurred: " + ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }


        #region checking for internet connection and writing text into richtextbox when is app loaded
        /// <summary>
        /// check if geocoding webs are online
        /// </summary>
        /// <param name="strServer">input url for API request</param>
        /// <returns></returns>
        public bool ConnectionAvailable(string strServer)
        {
            try
            {
                HttpWebRequest reqFp = (HttpWebRequest)WebRequest.Create(strServer);
                reqFp.Timeout = 5000;
                HttpWebResponse rspFp = (HttpWebResponse)reqFp.GetResponse();
                if (HttpStatusCode.OK == rspFp.StatusCode)
                {
                    // HTTP = 200 - Internet connection available, server online
                    rspFp.Close();
                    return true;
                }
                else
                {
                    // Other status - Server or connection not available
                    rspFp.Close();
                    return false;
                }
            }
            catch (WebException)
            {
                // Exception - connection not available

                return false;
            }
        }
        /// <summary>
        /// check for internet connection
        /// </summary>
        /// <returns>true if connection is on, false if not</returns>
        /// 
        public static bool CheckForInternetConnection()
        {
            try
            {
                using (var client = new WebClient())
                // ReSharper disable once UnusedVariable
                using (var stream = client.OpenRead("http://www.google.com"))
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            await Task.Factory.StartNew(CheckInternet);
            //Thread startReset = new Thread(CheckInternet);
            //startReset.Start();
        }
        public void CheckInternet()
        {
            Dispatcher.Invoke(DispatcherPriority.Background, new Action(() =>
            //delegate
            {
                try
                {
                    if (CheckForInternetConnection() == false)
                    {
                        MenuOpen.IsEnabled = false;
                        RtbProgress.AppendText("You don't have an Internet connection.\r");
                        RtbProgress.AppendText("Please refresh the connection (the program will run again) or quit the program.\rRESTARTING IN\r");
                        DoEvents();
                        for (int i = 11; i-- > 1;)
                        {
                            if (_quitClicked == true)
                            {
                                MItemQuit.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
                            }
                            RtbProgress.AppendText(i.ToString() + "\r");
                            Thread.Sleep(1000);
                            DoEvents();
                        }

                        if (_quitClicked == false)
                        {
                            if (MessageBox.Show("Do you want close the application? \rIf not, application will be restarted.", "No internet connection!", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
                            {
                                MItemReset.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
                            }
                            else
                            {
                                MItemQuit.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
                            }

                        }
                    }
                    else
                    {
                        // write this into richtextbox when is app successfully loaded
                        RtbProgress.AppendText("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\r");
                        RtbProgress.AppendText("Please CLOSE ALL Excel workbooks before geocoding.\rIf you have an open Excel during the process, errors can occur.\r");
                        RtbProgress.AppendText("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\r");
                        RtbProgress.AppendText("Waiting for .xlsx to be loaded\r");
                        RtbProgress.AppendText("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\r");
                        RtbProgress.ScrollToEnd();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("A handled exception just occurred: " + ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }));

        }
        #endregion

        /// <summary>
        /// check if APIs webs are online and functional
        /// </summary>
        private void CheckIfApisAreOnline()
        {
            // TODO: maybe this covert to keys, when they are right and check connection then
            Dictionary<object, string> urlsDic = new Dictionary<object, string>()
                            {
                                {CbxRuain, "http://www.vugtk.cz/euradin/services/rest.py/Geocode/text?SearchText=" },
                                {CbxGm, "https://maps.googleapis.com/maps/api/geocode/xml?address=" },
                                {CbxHm, "http://geocoder.cit.api.here.com/6.2/geocode.xml?app_id=O5Xux7fAgmj4kSi67XbA&app_code=cbeXpCRSVWo9kc17HbtHEA&gen=9&searchtext=Ostrava" },
                                {CbxMq, "http://open.mapquestapi.com/geocoding/v1/address?key=2i74YRgMWpE5GJoOlkFpy57yINNjQQ1V&outFormat=xml&location=Ostrava" },
                                {CbxMcz, "http://api.mapy.cz/geocode?query=" },
                                {CbxOsm, "http://nominatim.openstreetmap.org/search/cz/?format=xml" },
                                {CbxBm, "http://dev.virtualearth.net/REST/v1/Locations?q=Ostrava&o=xml&key=wrIA0ucuQwsQUxP6OAZP~GcpFgalVzrLfG6E-qgBnaQ~AhcLmtVfz7TlpdyF12sCtimnxcho0RXl_eW_FRJIYlDjaiGiq-a1lc2cZOxmKBIb" }
                            };
            foreach (var pair in urlsDic)
            {
                CheckBox cbx = pair.Key as CheckBox;
                if (ConnectionAvailable(pair.Value) == true)
                {
                    // ReSharper disable once PossibleNullReferenceException
                    RtbProgress.AppendText(cbx.Content + " geocoding web is ONLINE.\r");
                }
                else
                {
                    // ReSharper disable once PossibleNullReferenceException
                    RtbProgress.AppendText(cbx.Content + " geocoding web is OFFLINE.\r");
                    cbx.IsEnabled = false;
                }
            }
        }

        /// <summary>
        /// check if HereMaps API codes are good to go
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtboxHmAppId_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((string.IsNullOrWhiteSpace(TxtboxHmAppId.Text) || TxtboxHmAppId.Text == "Fill this key please!")
                &&
                (string.IsNullOrWhiteSpace(TxtboxHmAppCode.Text) || TxtboxHmAppCode.Text == "Fill this key please!"))
            {
                TxtboxHmAppId.Text = "Fill this key please!";
                TxtboxHmAppCode.Text = "Fill this key please!";
                ImgHm.Source = new BitmapImage(new Uri("/images/question.png", UriKind.Relative));
                CbxHm.IsEnabled = false;
            }
            else if ((string.IsNullOrWhiteSpace(TxtboxHmAppId.Text) || TxtboxHmAppId.Text == "Fill this key please!")
                    &&
                    (string.IsNullOrWhiteSpace(TxtboxHmAppCode.Text) == false || TxtboxHmAppCode.Text != "Fill this key please!"))
            {
                TxtboxHmAppId.Text = "Fill this key please!";
                ImgHm.Source = new BitmapImage(new Uri("/images/cancel.png", UriKind.Relative));
                CbxHm.IsEnabled = false;
            }
            else if ((string.IsNullOrWhiteSpace(TxtboxHmAppId.Text) == false || TxtboxHmAppId.Text != "Fill this key please!")
                    &&
                    (string.IsNullOrWhiteSpace(TxtboxHmAppCode.Text) || TxtboxHmAppCode.Text == "Fill this key please!"))
            {
                TxtboxHmAppCode.Text = "Fill this key please!";
                ImgHm.Source = new BitmapImage(new Uri("/images/cancel.png", UriKind.Relative));
                CbxHm.IsEnabled = false;
            }
            else
            {
                try
                {
                    string url = "http://geocoder.cit.api.here.com/6.2/geocode.xml?app_id=" + TxtboxHmAppId.Text + "&app_code=" + TxtboxHmAppCode.Text + "&gen=9&searchtext=Ostrava";
                    WebDownload wc = new WebDownload();
                    string s = wc.DownloadString(url);
                    if (RegexBetween(s, "<ViewId>(.*)</ViewId>") == "0")
                    {
                        ImgHm.Source = new BitmapImage(new Uri("/images/accept.png", UriKind.Relative));
                        CbxHm.IsEnabled = true;
                    }
                    else
                    {
                        ImgHm.Source = new BitmapImage(new Uri("/images/cancel.png", UriKind.Relative));
                        CbxHm.IsEnabled = false;
                    }
                }
                catch (WebException)
                {
                    ImgHm.Source = new BitmapImage(new Uri("/images/cancel.png", UriKind.Relative));
                    CbxHm.IsEnabled = false;
                }
            }
        }
        private void TxtboxHmAppCode_LostFocus(object sender, RoutedEventArgs e)
        {
            TxtboxHmAppId.RaiseEvent(new RoutedEventArgs(LostFocusEvent));
        }
        /// <summary>
        /// check if openMapQuest API codes are good to go
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtboxMqKey_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtboxMqKey.Text) || TxtboxMqKey.Text == "Fill this key please!")
            {
                TxtboxMqKey.Text = "Fill this key please!";
                ImgMq.Source = new BitmapImage(new Uri("/images/question.png", UriKind.Relative));
                CbxMq.IsEnabled = false;
            }
            else
            {
                try
                {
                    string url = "http://open.mapquestapi.com/geocoding/v1/address?key=" + TxtboxMqKey.Text +
                                 "&outFormat=xml&maxResults=1&thumbMaps=false&location=Ostrava";
                    WebDownload wc = new WebDownload();
                    string s = wc.DownloadString(url);
                    if (RegexBetween(s, "<statusCode>(.*)</statusCode>") == "0")
                    {
                        ImgMq.Source = new BitmapImage(new Uri("/images/accept.png", UriKind.Relative));
                        CbxMq.IsEnabled = true;
                    }
                    else
                    {
                        ImgMq.Source = new BitmapImage(new Uri("/images/cancel.png", UriKind.Relative));
                        CbxMq.IsEnabled = false;
                    }
                }
                catch (WebException)
                {
                    ImgMq.Source = new BitmapImage(new Uri("/images/cancel.png", UriKind.Relative));
                    CbxMq.IsEnabled = false;
                }
            }
        }
        /// <summary>
        /// check if GoogleMaps API codes are good to go
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtboxGmKey_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtboxGmKey.Text) || TxtboxGmKey.Text == "Fill this key please!")
            {
                TxtboxGmKey.Text = "Fill this key please!";
                ImgGm.Source = new BitmapImage(new Uri("/images/question.png", UriKind.Relative));
                CbxGm.IsEnabled = false;
            }
            else
            {
                try
                {
                    string url = "https://maps.googleapis.com/maps/api/geocode/xml?address=Ostrava&key=" + TxtboxGmKey.Text;
                    WebDownload wc = new WebDownload();
                    string s = wc.DownloadString(url);
                    if (RegexBetween(s, "<status>(.*)</status>") == "OK")
                    {
                        ImgGm.Source = new BitmapImage(new Uri("/images/accept.png", UriKind.Relative));
                        CbxGm.IsEnabled = true;
                    }
                    else
                    {
                        ImgGm.Source = new BitmapImage(new Uri("/images/cancel.png", UriKind.Relative));
                        CbxGm.IsEnabled = false;
                    }
                }
                catch (WebException)
                {
                    ImgGm.Source = new BitmapImage(new Uri("/images/cancel.png", UriKind.Relative));
                    CbxGm.IsEnabled = false;
                }
            }
        }
        /// <summary>
        /// check if BingMaps API codes are good to go
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtboxBmKey_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtboxBmKey.Text) || TxtboxBmKey.Text == "Fill this key please!")
            {
                TxtboxBmKey.Text = "Fill this key please!";
                ImgBm.Source = new BitmapImage(new Uri("/images/question.png", UriKind.Relative));
                CbxBm.IsEnabled = false;
            }
            else
            {
                try
                {
                    string url = "http://dev.virtualearth.net/REST/v1/Locations?q=Ostrava&o=xml&maxRes=1&key=" + TxtboxBmKey.Text;
                    WebDownload wc = new WebDownload();
                    string s = wc.DownloadString(url);
                    if (RegexBetween(s, "<StatusDescription>(.*)</StatusDescription>") == "OK" &&
                        RegexBetween(s, "<EstimatedTotal>(.*)</EstimatedTotal>") == "1")
                    {
                        ImgBm.Source = new BitmapImage(new Uri("/images/accept.png", UriKind.Relative));
                        CbxBm.IsEnabled = true;
                    }
                    else
                    {
                        ImgBm.Source = new BitmapImage(new Uri("/images/cancel.png", UriKind.Relative));
                        CbxBm.IsEnabled = false;
                    }
                }
                catch (WebException)
                {
                    ImgBm.Source = new BitmapImage(new Uri("/images/cancel.png", UriKind.Relative));
                    CbxBm.IsEnabled = false;
                }
            }
        }
    }
    public static class ExtensionMethods
    {
        private static readonly Action EmptyDelegate = delegate { };

        public static void Refresh(this UIElement uiElement)
        {
            uiElement.Dispatcher.Invoke(DispatcherPriority.Render, EmptyDelegate);
        }
    }
    public class WebDownload : WebClient
    {
        /// <summary>
        /// Time in milliseconds
        /// </summary>
        public int Timeout { get; set; }

        public WebDownload() : this(10000) { }

        public WebDownload(int timeout)
        {
            // ReSharper disable once ArrangeThisQualifier
            this.Timeout = timeout;
        }

        protected override WebRequest GetWebRequest(Uri address)
        {
            var request = base.GetWebRequest(address);
            if (request != null)
            {
                request.Timeout = Timeout;
            }
            return request;
        }
    }
}
