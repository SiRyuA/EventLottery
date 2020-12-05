using System;
using System.Timers;
using System.Windows;
using System.Windows.Threading;

namespace EventLottery
{
    /// <summary>
    /// Counting.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class Counting : Window
    {

        // 카운트 및 타이머
        public static int Count = 10;
        Timer timer = new Timer();

        public Counting()
        {
            InitializeComponent();
        }

        // 페이지 로딩 시 타이머 시작
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Count = 10;
            timer.Interval = 1000;
            timer.Elapsed += timerTick;
            timer.Enabled = true;
        }

        // 타이머 인터럽트에 따라 1초 마다 카운트 감소
        void timerTick(object sender, ElapsedEventArgs e)
        {
            Console.WriteLine(Count);

            if (Count < 1) // 0이 되면 타이머 종료 및 페이지 닫기
            {
                Application.Current.Dispatcher.Invoke(DispatcherPriority.Render, new Action(delegate
                {
                    timer.Stop();
                    Window.GetWindow(this).Hide();
                }));
            }

            // 출력
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Render, new Action(delegate
            {
                TB.Text = Count.ToString();
            }));

            Count--;
        }

    }
}
