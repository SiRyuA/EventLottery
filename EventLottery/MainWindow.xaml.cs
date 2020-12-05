using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace EventLottery
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        // 상품 목록 데이터 바인딩
        class ListExcel
        {
            public string value1 { get; set; }
            public string value2 { get; set; }
        }

        // 참가자 목록 데이터 바인딩
        class ListUser
        {
            public string Name { get; set; }
            public string Id { get; set; }
            public string Over { get; set; }
            public string Per { get; set; }
        }

        // 당첨자 목록 데이터 바인딩
        class ListWin
        {
            public string Item { get; set; }
            public string Name { get; set; }
            public string Id { get; set; }
        }

        // 임시 데이터
        List<ListExcel> listItems = new List<ListExcel>();
        List<ListUser> listUsers = new List<ListUser>();
        List<ListUser> listUsersBlind = new List<ListUser>();
        List<ListWin> listWins = new List<ListWin>();
        List<ListWin> listWinsBlind = new List<ListWin>();

        // 페이지
        Loading loading = new Loading();
        Counting counting = new Counting();

        public MainWindow()
        {
            InitializeComponent();
        }

        // 종료 될 경우 모든 페이지 다 종료
        private void Window_Closing(object sender, CancelEventArgs e)
        {
            counting.Close();
            loading.Close();
        }

        // 상품 목록 가져오기 클릭
        private void BtnExcelGet_Click(object sender, RoutedEventArgs e)
        {
            BackgroundWorker worker = new BackgroundWorker();
            worker.DoWork += workerGetList;
            worker.RunWorkerCompleted += workerCompleted;
            worker.RunWorkerAsync();
            loading.Show();
        }

        // 목록 추가 양식 클릭
        private void BtnGetExcel_Click(object sender, RoutedEventArgs e)
        {
            BackgroundWorker worker = new BackgroundWorker();
            worker.DoWork += workerGetForm;
            worker.RunWorkerCompleted += workerCompleted;
            worker.RunWorkerAsync();
            loading.Show();
        }

        // 당첨자 추첨 클릭
        private void BtnWinner_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("당첨자를 추첨합니다.", "추첨 시작", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                counting.Show();
                getWinner();
            }
            else
            {
                MessageBox.Show("추첨을 취소합니다.");
            }
        }

        // 당첨자 목록 엑셀 저장
        private void BtnWinExcel_Click(object sender, RoutedEventArgs e)
        {
            BackgroundWorker worker = new BackgroundWorker();
            worker.DoWork += workerWinner;
            worker.RunWorkerCompleted += workerCompleted;
            worker.RunWorkerAsync();
        }

        // 백그라운드 워커 - 목록 가져오기
        void workerGetList(object sender, DoWorkEventArgs e)
        {
            getList();
            MessageBox.Show("목록을 가져왔습니다.");
        }

        // 백그라운드 워커 - 양식 다운로드
        void workerGetForm(object sender, DoWorkEventArgs e)
        {
            GetExcelForm();
            MessageBox.Show("목록 양식을 저장했습니다.");
        }

        // 백그라운드 워커 - 당첨자 목록 저장
        void workerWinner(object sender, DoWorkEventArgs e)
        {
            WinToExcel();
            MessageBox.Show("당첨자 목록을 저장했습니다.");
        }

        // 백그라운드 워커 종료 시 로딩 창 닫기
        void workerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            loading.Hide();
        }

        // 목록 가져오기 처리
        private void getList()
        {
            string path = GetFilePath(); // 경로 가져오기
            getItemList(path);
            getUserList(path);
        }

        // 상품 가져오기
        private void getItemList(string path)
        {
            if (path.Length > 0)
            {
                listItems = ExcelToList(path, 1);

                if (listItems.Count > 0) // 행이 있을 경우 -> 데이터가 있을 경우
                {
                    Application.Current.Dispatcher.Invoke(DispatcherPriority.Render, new Action(delegate
                    {
                        listviewItem.ItemsSource = listItems;
                        listviewItem.Items.Refresh();
                    }));
                }
                else
                {
                    MessageBox.Show("데이터가 없습니다.");
                }
            }
        }

        // 참가자 가져오기
        private void getUserList(string path)
        {
            if (path.Length > 0) // 경로 값이 있을 경우
            {
                // 데이터 처리 목적 리스트 생성
                List<ListExcel> list = new List<ListExcel>();
                List<string> pass = new List<string>();

                // 엑셀 데이터 가져오기
                list = ExcelToList(path, 2);
                int lCnt = list.Count;

                if (lCnt > 0) // 행이 있을 경우 -> 데이터가 있을 경우
                {
                    listUsers.Clear(); // 초기화
                    listUsersBlind.Clear();

                    for (int i = 0; i < lCnt; i++) // 집계
                    {
                        string id1 = list[i].value2;
                        int cnt = 0;

                        if (!pass.Contains(id1)) // 중복 등록 목록에 없을 경우
                        {
                            for (int j = 0; j < lCnt; j++) // 동일한 값 찾아서 중복 등록
                            {
                                string id2 = list[j].value2;

                                if (id1 == id2)
                                {
                                    cnt++;
                                    pass.Add(id2);
                                }
                            }

                            // 데이터 처리
                            string name = list[i].value1;
                            string over = cnt.ToString();
                            double x = (double)cnt / (double)lCnt;
                            string per = (x * 100.0).ToString() + "%";

                            string idBlind = id1.Substring(0, 4) + "****";

                            // 참가자 목록 리스트에 추가
                            listUsers.Add(new ListUser() { Name = name, Id = id1, Over = over, Per = per });
                            listUsersBlind.Add(new ListUser() { Name = name, Id = idBlind, Over = over, Per = per });
                        }
                    }

                    Application.Current.Dispatcher.Invoke(DispatcherPriority.Render, new Action(delegate
                    {
                        // 참가자 목록 리스트뷰 갱신
                        //listviewUser.Items.Clear();
                        listviewUser.ItemsSource = listUsersBlind;
                        listviewUser.Items.Refresh();
                    }));
                }
                else
                {
                    MessageBox.Show("데이터가 없습니다.");
                }
            }
        }

        // 파일 경로 가져오기
        private string GetFilePath()
        {
            // 다이얼로그 오픈
            OpenFileDialog file = new OpenFileDialog();
            file.Filter = "Excel (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls";

            // 경로 리턴
            if (file.ShowDialog() == false)
            {
                return "";
            }
            else
            {
                return file.FileName;
            }
        }

        // 엑셀 파일 열어서 리스트 보관 처리
        private List<ListExcel> ExcelToList(string path, int target)
        {
            // 데이터 처리 목적 리스트 생성
            var list = new List<ListExcel>();

            try
            {
                var app = new Excel.Application();
                var book = app.Workbooks.Open(path, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                var sheet = (Excel.Worksheet)book.Worksheets.get_Item(target);

                app.Visible = false;

                // 시트 행 값 가져오기
                var range = sheet.UsedRange;
                var row = range.Rows.Count;

                // 행 마다 데이터 리스트 저장
                for (int i = 2; i <= row; i++)
                {
                    string v1 = ((range.Cells[i, 1] as Excel.Range).Value2).ToString();
                    string v2 = ((range.Cells[i, 2] as Excel.Range).Value2).ToString();

                    list.Add(new ListExcel() { value1 = v1, value2 = v2 });
                }

                // 닫기
                book.Close(false, null, null);
                app.Quit();

                releaseObject(sheet);
                releaseObject(book);
                releaseObject(app);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            return list;
        }

        // 엑셀로 당첨자 목록 저장
        private void WinToExcel()
        {
            try
            {
                // 경로 생성
                string path = saveFilePath();
                if (path.Length < 1) return;

                // 엑셀 실행
                var app = new Excel.Application();
                var book = app.Workbooks.Add();
                var sheet = (Excel.Worksheet)book.Worksheets.get_Item(1);

                // 데이터 입력
                sheet.Cells[1, 1] = "상품";
                sheet.Cells[1, 2] = "닉네임";
                sheet.Cells[1, 3] = "ID";

                for(int i = 0; i < listWins.Count; i++)
                {
                    sheet.Cells[2 + i, 1] = listWins[i].Item;
                    sheet.Cells[2 + i, 2] = listWins[i].Name;
                    sheet.Cells[2 + i, 3] = listWins[i].Id;
                }

                // 너비 조절
                sheet.Columns.AutoFit();

                // 저장 후 종료
                book.SaveAs(path, Excel.XlFileFormat.xlWorkbookDefault);
                book.Close(true);
                app.Quit();

                releaseObject(sheet);
                releaseObject(book);
                releaseObject(app);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        // 목록 추가 양식 저장
        private void GetExcelForm()
        {
            try
            {
                // 경로 생성
                string path = saveFilePath();
                if (path.Length < 1) return;

                // 엑셀 실행
                var app = new Excel.Application();
                var book = app.Workbooks.Add();

                // 시트 1번
                var sheet = (Excel.Worksheet)book.Worksheets.get_Item(1);
                
                // 데이터 입력
                sheet.Cells[1, 1] = "상품";
                sheet.Cells[1, 2] = "개수";

                var sheet2 = book.Worksheets.Add();

                // 데이터 입력
                sheet2.Cells[1, 1] = "닉네임";
                sheet2.Cells[1, 2] = "ID";

                sheet2.Move(Missing.Value, book.Sheets[book.Sheets.Count]);

                // 저장 후 종료
                book.SaveAs(path, Excel.XlFileFormat.xlWorkbookDefault);
                book.Close(true);
                app.Quit();

                releaseObject(sheet2);
                releaseObject(sheet);
                releaseObject(book);
                releaseObject(app);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        // 메모리 반환
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        // 저장 경로 지정하기
        private string saveFilePath()
        {
            // 경로 생성
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            saveFile.Title = "저장 경로를 지정하세요.";
            saveFile.OverwritePrompt = true;
            saveFile.DefaultExt = "xlsx";
            saveFile.Filter = "Excel (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls";

            // 저장 버튼 누를 경우 저장 경로 반환
            if (saveFile.ShowDialog() == true)
            {
                return saveFile.FileName;
            }
            else
            {
                return "";
            }
        }

        // 당첨자 생성
        private void getWinner()
        {
            // 당첨자 초기화 및 리스트뷰 갱신
            listWins.Clear();
            listWinsBlind.Clear();

            // 중복 당첨 방지 목적 리스트 생성
            List<string> pass = new List<string>();

            // 전체 상품 개수 측정 (중복 추첨 여부 판단 목적)
            int itemAll = 0;
            for (int i = 0; i < listItems.Count; i++)
            {
                itemAll += Convert.ToInt32(listItems[i].value2);
            }

            // 전체 유저 수
            int userAll = listUsers.Count;

            // 전체 상품 개수만큼 반복
            for (int i = 0; i < listItems.Count; i++)
            {
                int cnt = Convert.ToInt32(listItems[i].value2);
                int loop = 0;

                while (loop < cnt)
                {
                    // 전체 참가자 수 범위 내 랜덤으로 숫자 뽑기
                    Random rand = new Random();
                    int r = rand.Next(listUsers.Count - 1);

                    string id = listUsers[r].Id; // 당첨자 ID

                    if (!pass.Contains(id) || itemAll > userAll)
                    {
                        string item = listItems[i].value1; // 상품 데이터
                        string name = listUsers[r].Name; // 당첨자 닉네임
                        string idBlind = listUsersBlind[r].Id;

                        listWins.Add(new ListWin() { Item = item, Name = name, Id = id });
                        listWinsBlind.Add(new ListWin() { Item = item, Name = name, Id = idBlind });
                        pass.Add(id);
                        loop++;
                    }
                }
            }

            // 당첨자 초기화 및 리스트뷰 갱신
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Render, new Action(delegate
            {
                listviewWin.ItemsSource = listWinsBlind;
                listviewWin.Items.Refresh();
            }));
            
        }

    }
}
