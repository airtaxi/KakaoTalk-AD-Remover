using KaTalkEspresso;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace KakaoTalk_AD_Remover
{
    public partial class MainWindow : Window
    {
        private readonly Logger log = Logger.getInstance();
        private IntPtr hWndKaTalk = IntPtr.Zero;
        private double step = 0;
        private bool isInForeground = true;
        private double tryCount = 0;
        private string ktPath = null;
        private bool isFirstRun = false;

        /// <summary>
        /// 프로세스 목록에서 지정된 경로의 카카오톡 프로세스를 찾아 메인 창 핸들을 반환합니다.
        /// </summary>
        /// <param name="predefinedKaTalkPath">카카오톡 경로</param>
        /// <returns>카카오톡 메인 창 핸들</returns>
        private IntPtr detectKakaoTalk(string predefinedKaTalkPath)
        {
            IntPtr handle = IntPtr.Zero;

            if (predefinedKaTalkPath == null || "".Equals(predefinedKaTalkPath))
            {
                // 경로가 비어있으면 진행 불가
                log.error("path not provided. cannot continue.");

                return handle;
            }

            log.info("checking for running KakaoTalk's path provided as -> " + predefinedKaTalkPath);

            FileInfo katalkExe = new FileInfo(predefinedKaTalkPath);

            if (!katalkExe.Exists)
            {
                // 경로의 파일이 존재하지 않습니다.
                // 진행 불가
                log.error("Could not find -> " + katalkExe.FullName);

                return handle;
            }


            string katalkExeStr = katalkExe.Name;
            int lastExeFound = katalkExeStr.ToLower().LastIndexOf(".exe");
            if (lastExeFound < 0)
            {
                // exe 확장명을 경로로부터 찾지 못함. 검색 불가
                log.warn("\".exe\" from name not found. it may cannot find for \"" + katalkExeStr + "\" from running processes.");

                return handle;
            }
            katalkExeStr = katalkExeStr.Substring(0, lastExeFound);

            //카톡 exe 이름은 소문자로 바꾸고 이걸 비교에 사용
            katalkExeStr = katalkExeStr.ToLower();

            // 이 경로로 된 프로세스가 있는지 찾기 위해 프로세스 목록 수집.
            Process[] procs = Process.GetProcesses();


            log.info("Checking for running " + procs.Length + " processes to pinpoint target executable.");
            foreach (Process proc in procs)
            {
                if (katalkExeStr.Equals(proc.ProcessName.ToLower()))//Util.toLowerComplyingFs(proc.ProcessName)))
                {
                    log.info("Found \"" + proc.ProcessName + "\" as candidate.");
                    //kakaotalk 프로세스 발견
                    // MainModule 속성 접근시 권한 문제 발생할 수 있으므로 catch
                    try
                    {
                        // 지정된 경로와 카카오톡 경로와 일치하는 경우
                        if (katalkExe.FullName.Equals(proc.MainModule.FileName))
                        {
                            log.info("this process matches with provided path. using this as further use.");
                            handle = proc.MainWindowHandle;//카톡 메인창 핸들
                                                           //handle = proc.Handle;// 카톡 핸들    
                                                           //MessageBox.Show("detection successful");
                        }
                    }
                    catch (Win32Exception w32e)
                    {
                        // Access Denied 문제 발생
                        log.error("Exception raised. Maybe could not access process info?");
                        log.error(w32e.ToString());
                    }
                }
            }
            return handle;
        }

        /// <summary>
        /// 카카오톡이 같은 폴더에 있으면  반환합니다.
        /// </summary>
        /// <returns>카톡 exe, 같은 폴더에 있으니 파일 이름만 있어도 된다</returns>
        private string openKakaoSameDir()
        {
            string exeOnSameDir = null;
            try
            {
                //같은 폴더에 카카오톡 실행파일이 있는지 확인
                string filename = "KakaoTalk.exe";
                FileInfo katalkExe = new FileInfo(filename);
                if (katalkExe.Exists)
                {
                    // 같은 폴더에 있음
                    exeOnSameDir = filename;
                }
                else
                {
                    //같은 경로에 없음
                    log.warn(filename + " not found on same dir");
                }
            }
            catch (Exception e)
            {
                log.warn("Exception raised during registry search for kakaoopen");
                log.warn(e.ToString());
            }
            return exeOnSameDir;
        }

        /// <summary>
        /// 레지스트리의 카카오톡 kakaoopen 키를 읽어 경로를 찾아냅니다.
        /// </summary>
        /// <returns>kakaoopen키 상의 카카오톡 경로</returns>
        private string openKakaoRegistryLocation()
        {
            // 추출한 경로가 기억될 변수
            string openKakaoReg = null;

            // 레지스트리를 닫아야 하므로 try 밖에서 선언
            RegistryKey kakaoOpen = null;
            try
            {
                // HKCR에서 카카오open 커맨드 레지스트리를 읽어 가져옵니다.
                kakaoOpen = Registry.ClassesRoot.OpenSubKey(@"kakaoopen\shell\open\command\", false);

                if (kakaoOpen == null)
                {
                    //레지스트리 값이 존재하지 않으면 검색할 필요가 없음.
                    log.warn("registry for kakaoopen not found");

                    return null;
                }
                // 레지스트리 (기본값)을 읽어옴
                //"C:\Program Files (x86)\Kakao\KakaoTalk\KakaoTalk.exe" "%1"
                openKakaoReg = (string)kakaoOpen.GetValue("");

                log.info("registry entry (default) found for kakaoopen, value -> " + openKakaoReg);

                bool exeExtracted = false;
                //\".*?\" 이전 원소 0 이상 최소한으로 일치.
                foreach (Match match in Regex.Matches(openKakaoReg, "\"([^\"]*)\""))
                {
                    // 앞뒤 따옴표 제거
                    string matchSub = match.ToString();

                    // subString시 startIndex의 글자는 포함되므로 + 1
                    int firstQuote = matchSub.IndexOf('\"') + 1;
                    // subString시 endIndex의 글자는 포함되지 않으므로 그대로 둠
                    int lastQuote = matchSub.LastIndexOf('\"');

                    if (firstQuote < 0 || firstQuote == lastQuote)
                    {
                        // 따옴표를 못 찾았습니다.
                        // 또는
                        // 마지막 따옴표가 첫 따옴표가 같으면 완전히 감싸진 모양이 아님

                        //다음 항목으로 넘어감
                        continue;
                    }
                    if (matchSub.ToLower().Contains(".exe"))
                    {
                        // 경로에 실행파일이 포함된 경우 따옴표 제거된 문자열을 알아둠
                        matchSub = matchSub.Substring(firstQuote, lastQuote - firstQuote);
                        openKakaoReg = matchSub;

                        exeExtracted = true;
                    }
                }

                if (exeExtracted)
                {
                    // 경로 추출 성공
                    log.info("file path from registry extracted -> " + openKakaoReg);
                }
                else
                {
                    // 경로 추출 실패
                    log.warn("failed to extract file path from registry");
                }

            }
            catch (Exception e)
            {
                log.warn("Exception raised during registry search for kakaoopen");
                log.warn(e.ToString());
            }
            finally
            {
                // 레지스트리를 접근했으면 닫아줘야 한다.
                if (kakaoOpen != null)
                {
                    try
                    {
                        log.info("closing registry handle for kakaoopen");

                        kakaoOpen.Close();
                    }
                    catch (Exception e)
                    {
                        log.warn("Exception raised during closing registry handle for kakaoopen");
                        log.warn(e.ToString());
                    }
                }
            }

            return openKakaoReg;
        }
        private bool runningKakaoTalk()
        {
            // 카톡 실행 대기 중
            // 카톡 핸들을 찾고 알아둡니다.
            hWndKaTalk = detectKakaoTalk(ktPath);
            // 핸들이 null이 아닌지 확인
            return !IntPtr.Zero.Equals(hWndKaTalk);
        }

        public async void doStep()
        {
            await Task.Delay(100);
            if (step == 0)
            {
                ktPath = openKakaoRegistryLocation();

                if (ktPath == null)
                {
                    //레지스트리에 카톡이 없으면 같은 폴더에서 카톡 검색
                    ktPath = openKakaoSameDir();
                }

                if (ktPath != null)
                {
                    // 카톡 경로 찾음
                    TB_Main.Text = "카카오톡 경로 발견됨";

                    step++;
                    doStep();
                }
                else
                {
                    // 카톡 경로 못 찾음
                    TB_Main.Text = "오류 : 카카오톡 설치를 확인하지 못했습니다.";
                }
            }
            else if (step == 1)
            {
                if (runningKakaoTalk())
                {
                    //카톡 핸들을 찾았다. 다음 단계로
                    TB_Main.Text = "카카오톡 실행 감지됨. 대기중";
                    step++;
                    doStep();
                }
                else
                {
                    try
                    {
                        TB_Main.Text = "실행된 카카오톡 프로세스 발견되지 않음. 대기중";
                        if(!isFirstRun)
                        {
                            isFirstRun = true;
                            ProcessStartInfo ktExePath = new ProcessStartInfo(ktPath);
                            Process.Start(ktExePath);
                            
                            await Task.Delay(3000);
                        }
                        doStep();
                    }
                    catch (Exception launchEx)
                    {
                        TB_Main.Text = "오류 : 카카오톡 실행에 실패했습니다.";
                        // 실행도중 예외 발생
                        log.warn("Cannot start " + ktPath);
                        log.warn(launchEx.ToString());
                    }
                }
            }
            else if (step < 3)
            {
                TB_Main.Text = $"광고를 숨기는중입니다 ({tryCount++})";
                int hideResult = 0;
                try
                {
                    hideResult = AdCloser.closeAdsKakaoTalk(hWndKaTalk);
                }
                catch (Exception) { }
                if (hideResult == 0)
                {
                    step += 0.5;
                    doStep();
                }
                else
                {
                    step = 0;
                    doStep();
                }
            }
            else
            {
                PB_Main.IsIndeterminate = false;
                PB_Main.Value = 100;
                TB_Main.Text = "완료. 3초뒤 백그라운드로 전환됩니다.";

                if (isInForeground)
                {
                    isInForeground = false;
                    await Task.Delay(3000);
                    Hide();
                }

                int hideResult = 1;
                try
                {
                    hideResult = AdCloser.closeAdsKakaoTalk(hWndKaTalk);
                }
                catch (Exception) { }
                finally
                {
                    if (hideResult != 0)
                    {
                        step = 0;
                    }
                }
                doStep();
            }
        }

        public MainWindow()
        {
            InitializeComponent();

            doStep();
        }
    }
}
