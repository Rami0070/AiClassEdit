using System;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Win32;

namespace SmartClassApp
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                // 1. 단독 실행 전용 디렉터리 준비 (%TEMP%\SmartClassApp)
                string tempDir = Path.Combine(Path.GetTempPath(), "SmartClassApp");
                if (!Directory.Exists(tempDir))
                {
                    Directory.CreateDirectory(tempDir);
                }

                string targetHtml = Path.Combine(tempDir, "index.html");
                string profileDir = Path.Combine(tempDir, "profile");

                // 2. exe 내부에 내장된 최신 index.html 리소스를 추출하여 항상 최신 버전으로 저장
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (Stream resourceStream = assembly.GetManifestResourceStream("AppIndexHtml"))
                {
                    if (resourceStream == null)
                    {
                        // 만약 내장 리소스가 없을 경우 exe와 동일한 디렉터리의 index.html fallback
                        string fallbackHtml = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "index.html");
                        if (File.Exists(fallbackHtml))
                        {
                            targetHtml = fallbackHtml;
                        }
                        else
                        {
                            MessageBox.Show(
                                "프로그램 내장 리소스를 불러올 수 없습니다.",
                                "스마트 학급편성 시스템",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error
                            );
                            return;
                        }
                    }
                    else
                    {
                        using (FileStream fileStream = new FileStream(targetHtml, FileMode.Create, FileAccess.Write))
                        {
                            byte[] buffer = new byte[81920];
                            int bytesRead;
                            while ((bytesRead = resourceStream.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                fileStream.Write(buffer, 0, bytesRead);
                            }
                        }
                    }
                }

                string fileUri = "file:///" + targetHtml.Replace('\\', '/');

                // 3. 브라우저 간섭 없는 독립 데스크톱 앱 창 모드(--app) 실행
                string browserPath = FindChromiumBrowser();

                if (!string.IsNullOrEmpty(browserPath) && File.Exists(browserPath))
                {
                    // 독립적인 전용 프로필 디렉터리를 지정하여 일반 브라우저(탭/주소창/북마크)와 완전히 분리된 전용 데스크톱 창으로 구동
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = browserPath,
                        Arguments = string.Format(
                            "--user-data-dir=\"{0}\" --app=\"{1}\" --window-size=1280,860 --no-first-run --no-default-browser-check --disable-features=Translate,OptimizationHints,MediaRouter",
                            profileDir,
                            fileUri
                        ),
                        UseShellExecute = false
                    };

                    Process proc = Process.Start(psi);
                    if (proc != null)
                    {
                        proc.WaitForExit();
                    }
                }
                else
                {
                    // Chromium 브라우저가 없는 예외적인 환경: 시스템 기본 웹 브라우저로 실행
                    ProcessStartInfo psi = new ProcessStartInfo
                    {
                        FileName = targetHtml,
                        UseShellExecute = true
                    };
                    Process.Start(psi);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "프로그램 실행 중 오류가 발생했습니다:\n" + ex.Message,
                    "스마트 학급편성 시스템",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        static string FindChromiumBrowser()
        {
            // 1. 레지스트리 App Paths 확인
            string[] regKeys = new string[]
            {
                @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe",
                @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
                @"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\whale.exe",
                @"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe",
                @"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
                @"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\whale.exe"
            };

            foreach (string key in regKeys)
            {
                try
                {
                    object val = Registry.GetValue(key, "", null);
                    if (val != null)
                    {
                        string path = val.ToString().Trim('\"');
                        if (File.Exists(path)) return path;
                    }
                }
                catch { }
            }

            // 2. 표준 설치 디렉터리 확인
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string progFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            string progFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);

            string[] directPaths = new string[]
            {
                Path.Combine(progFilesX86, @"Microsoft\Edge\Application\msedge.exe"),
                Path.Combine(progFiles, @"Microsoft\Edge\Application\msedge.exe"),
                Path.Combine(localAppData, @"Microsoft\Edge\Application\msedge.exe"),
                Path.Combine(progFiles, @"Google\Chrome\Application\chrome.exe"),
                Path.Combine(progFilesX86, @"Google\Chrome\Application\chrome.exe"),
                Path.Combine(localAppData, @"Google\Chrome\Application\chrome.exe"),
                Path.Combine(progFiles, @"Naver\Naver Whale\Application\whale.exe"),
                Path.Combine(progFilesX86, @"Naver\Naver Whale\Application\whale.exe")
            };

            foreach (string p in directPaths)
            {
                if (File.Exists(p)) return p;
            }

            return null;
        }
    }
}
