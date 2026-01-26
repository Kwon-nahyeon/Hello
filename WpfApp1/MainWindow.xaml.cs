using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

using Dicom;
using Dicom.Imaging;

// 폴더 / 파일 선택창
using Forms = System.Windows.Forms;

namespace WpfApp1
{
    
    // DICOM 비식별화 규칙 모음
       public static class DicomAnonymizer
    {
       
        /// 전달받은 DicomDataset에 비식별화 규칙을 적용
        
        public static void AnonymizeDataset(DicomDataset ds)
        {
            if (ds == null) return;

            // ===== 1) 삭제 대상 =====
            // (0010,0010) 환자 이름
            RemoveIfExists(ds, DicomTag.PatientName);

            // (0010,0020) 환자 ID
            RemoveIfExists(ds, DicomTag.PatientID);

            // (0010,0030) 환자 생년월일
            RemoveIfExists(ds, DicomTag.PatientBirthDate);

            // (0010,0040) 환자 성별 → 유지 (아무 것도 안 함)

            // (0010,1000~109C) 기타 환자 ID
            RemovePatientOtherIDs(ds);

            // 날짜/시간 관련
            RemoveIfExists(ds, DicomTag.StudyDate);          // (0008,0020)
            RemoveIfExists(ds, DicomTag.StudyTime);          // (0008,0030)
            RemoveIfExists(ds, DicomTag.SeriesDate);         // (0008,0021)
            RemoveIfExists(ds, DicomTag.AcquisitionDate);    // (0008,0022)

            // 두 번째 표의 삭제 대상
            RemoveIfExists(ds, DicomTag.EthnicGroup);            // (0010,2160) 인종
            RemoveIfExists(ds, DicomTag.Occupation);             // (0010,2180) 직업
            RemoveIfExists(ds, DicomTag.InstitutionName);        // (0008,0080) 기관명
            RemoveIfExists(ds, DicomTag.InstitutionAddress);     // (0008,0081) 기관 주소
            RemoveIfExists(ds, DicomTag.ReferringPhysicianName); // (0008,0090) 참조의 이름
            RemoveIfExists(ds, DicomTag.PerformingPhysicianName);// (0008,1050) 수행의
            RemoveIfExists(ds, DicomTag.OperatorsName);          // (0008,1070) 운영자 이름

            // ===== 2) 유지 대상 =====
            // PatientSex(0010,0040), PatientAge(0010,1010),
            // PatientSize(0010,1020), PatientWeight(0010,1030) 등은 유지 → 건드리지 않음

            // ===== 3) UID 재생성 =====
            RegenerateUid(ds, DicomTag.StudyInstanceUID);   // (0020,000D)
            RegenerateUid(ds, DicomTag.SeriesInstanceUID);  // (0020,000E)
            RegenerateUid(ds, DicomTag.SOPInstanceUID);     // (0008,0018)
        }

        private static void RemoveIfExists(DicomDataset ds, DicomTag tag)
        {
            if (ds.Contains(tag))
            {
                ds.Remove(tag);
            }
        }

      
        /// (0010,1000 ~ 0010,109C) 범위의 기타 환자 ID 태그 삭제
        private static void RemovePatientOtherIDs(DicomDataset ds)
        {
            var toRemove = ds
                .Where(x => x.Tag.Group == 0x0010 &&
                            x.Tag.Element >= 0x1000 &&
                            x.Tag.Element <= 0x109C)
                .ToList();

            foreach (var item in toRemove)
            {
                ds.Remove(item.Tag);
            }
        }

        private static void RegenerateUid(DicomDataset ds, DicomTag tag)
        {
            if (!ds.Contains(tag)) return;

            var newUid = DicomUIDGenerator.GenerateDerivedFromUUID();
            ds.AddOrUpdate(tag, newUid.UID);
        }
    }

    // ListView에 뿌릴 Key/Value 아이템
    public class TagItem
    {
        public string Key { get; set; }
        public string Value { get; set; }
    }

    public partial class MainWindow : Window
    {
        // 태그 표시용 컬렉션
        public ObservableCollection<TagItem> TagItems { get; } = new ObservableCollection<TagItem>();

        private string[] _files = Array.Empty<string>();
        private int _index = 0;

        // 스크롤바 값 바꿀 때 이벤트 중복 호출 방지
        private bool _suppressScrollEvent = false;

        // PixelData(영상) 없는 파일 경고를,
        // 같은 파일에 대해서는 한 번만 띄우기 위한 기록
        private string _lastPixelDataWarningFile = null;

        public MainWindow()
        {
            InitializeComponent();
            Loaded += MainWindow_Loaded;

            // DataContext를 윈도우 자신으로 잡아 바인딩 가능하게
            DataContext = this;
        }

        /// 하단 로그창(ListBox)에 한 줄 출력
        private void AddLog(string level, string message)
        {
            if (LogList == null) return;

            string line = $"[{DateTime.Now:HH:mm:ss}] [{level}] {message}";
            LogList.Items.Add(line);
            LogList.ScrollIntoView(line);
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // 처음 실행 시 폴더 한 번 선택
            string folder = PickDicomFolder();
            if (string.IsNullOrEmpty(folder))
            {
                Close();
                return;
            }

            LoadFolder(folder);

            if (_files.Length == 0)
            {
                MessageBox.Show("선택한 폴더에 .dcm 파일이 없습니다.", "알림");
                Close();
                return;
            }

            // 중간부터 시작
            _index = _files.Length / 2;

            // 스크롤바 범위 세팅
            SliceScrollBar.Minimum = 0;
            SliceScrollBar.Maximum = _files.Length - 1;

            _suppressScrollEvent = true;
            SliceScrollBar.Value = _index;
            _suppressScrollEvent = false;

            ShowDicom(_files[_index]);
        }

        // ==========================
        // 상단 버튼 이벤트 핸들러
        // ==========================

        // 파일 한 장 선택해서 보기
        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new Forms.OpenFileDialog())
            {
                dlg.Filter = "DICOM 파일 (*.dcm)|*.dcm|모든 파일 (*.*)|*.*";
                dlg.Title = "DICOM 파일 열기";

                if (dlg.ShowDialog() == Forms.DialogResult.OK)
                {
                    _files = new[] { dlg.FileName };
                    _index = 0;

                    // 스크롤바 세팅 (한 장이라도 동일하게 처리)
                    SliceScrollBar.Minimum = 0;
                    SliceScrollBar.Maximum = _files.Length - 1;

                    _suppressScrollEvent = true;
                    SliceScrollBar.Value = _index;
                    _suppressScrollEvent = false;

                    _lastPixelDataWarningFile = null; // 새 파일이니 경고 기록 초기화
                    AddLog("INFO", $"파일 열기: {dlg.FileName}");
                    ShowDicom(_files[_index]);
                }
            }
        }

        // 폴더 선택해서 그 안의 .dcm들을 모두 로드
        private void OpenFolder_Click(object sender, RoutedEventArgs e)
        {
            string folder = PickDicomFolder();
            if (string.IsNullOrEmpty(folder)) return;

            LoadFolder(folder);

            if (_files.Length == 0)
            {
                MessageBox.Show("선택한 폴더에 .dcm 파일이 없습니다.", "알림");
                return;
            }

            _index = _files.Length / 2;

            SliceScrollBar.Minimum = 0;
            SliceScrollBar.Maximum = _files.Length - 1;

            _suppressScrollEvent = true;
            SliceScrollBar.Value = _index;
            _suppressScrollEvent = false;

            _lastPixelDataWarningFile = null;
            AddLog("INFO", $"폴더 열기: {folder}");
            ShowDicom(_files[_index]);
        }

        /// 현재 보고 있는 DICOM 1장을 비식별화해서
        /// 원본 폴더 하위 "Anonymized" 폴더에 저장

        private void Anonymize_Click(object sender, RoutedEventArgs e)
        {
            if (_files == null || _files.Length == 0)
            {
                MessageBox.Show("먼저 폴더를 열어 DICOM 이미지를 로드한 뒤 사용하세요.", "알림");
                return;
            }

            string inputPath = _files[_index];

            try
            {
                // DICOM 파일 열기
                var dicomFile = DicomFile.Open(inputPath);

                // 비식별화 규칙 적용
                DicomAnonymizer.AnonymizeDataset(dicomFile.Dataset);

                // 저장 폴더: 원본 폴더 아래 Anonymized 폴더 생성
                string dir = Path.GetDirectoryName(inputPath);
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
                string ext = Path.GetExtension(inputPath);

                string outDir = Path.Combine(dir, "Anonymized");
                Directory.CreateDirectory(outDir);

                // 파일명 뒤에 _anon 붙여서 저장
                string outputPath = Path.Combine(outDir, fileNameWithoutExt + "_anon" + ext);

                dicomFile.Save(outputPath);

                AddLog("INFO", $"비식별화 저장 완료: {outputPath}");
                MessageBox.Show(
                    "비식별화된 DICOM 파일이 저장되었습니다.\n\n" +
                    $"경로: {outputPath}",
                    "비식별 저장 완료",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AddLog("ERROR", $"비식별화 실패: {ex.Message}");
                MessageBox.Show(ex.ToString(), "비식별화 실패");
            }
        }

        /// <summary>
        /// AQ 버튼: 현재 화면에 보이는 DICOM 1장에 대해
        ///  - 비식별 태그 삭제 여부
        ///  - 필수 태그 존재 여부
        ///  - 몇 개 수치 범위
        ///  - UID 형식
        /// 를 검사
        /// </summary>
        private void RunAQ_Click(object sender, RoutedEventArgs e)
        {
            if (_files == null || _files.Length == 0)
            {
                MessageBox.Show("먼저 DICOM 파일을 열어주세요.", "AQ", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string path = _files[_index];

            try
            {
                var dicomFile = DicomFile.Open(path);
                var ds = dicomFile.Dataset;

                var sb = new StringBuilder();
                bool allPass = true;

                sb.AppendLine($"[파일] {Path.GetFileName(path)}");

                // 1) 비식별화로 제거되어야 하는 태그들 확인
                allPass &= AqCheckRemovedTag(ds, DicomTag.PatientName, "Patient's Name", sb);
                allPass &= AqCheckRemovedTag(ds, DicomTag.PatientID, "Patient ID", sb);
                allPass &= AqCheckRemovedTag(ds, DicomTag.PatientBirthDate, "Patient's Birth Date", sb);
                allPass &= AqCheckRemovedTag(ds, DicomTag.ReferringPhysicianName, "Referring Physician's Name", sb);
                allPass &= AqCheckRemovedTag(ds, DicomTag.InstitutionName, "Institution Name", sb);

                // 2) 유지/필수 태그 존재 여부
                allPass &= AqCheckRequiredTagString(ds, DicomTag.PatientSex, "Patient's Sex", sb);
                allPass &= AqCheckRequiredTagString(ds, DicomTag.Modality, "Modality", sb);
                allPass &= AqCheckRequiredTagString(ds, DicomTag.StudyInstanceUID, "Study Instance UID", sb);
                allPass &= AqCheckRequiredTagString(ds, DicomTag.SeriesInstanceUID, "Series Instance UID", sb);
                allPass &= AqCheckRequiredTagString(ds, DicomTag.SOPInstanceUID, "SOP Instance UID", sb);

                // 3) 수치 범위 간단 체크 (값이 있을 때만 검사)
                allPass &= AqCheckNumericRange(ds, DicomTag.KVP, "KVP", 50, 150, sb);
                allPass &= AqCheckNumericRange(ds, DicomTag.ExposureTime, "Exposure Time", 1, 3000, sb);
                allPass &= AqCheckNumericRange(ds, DicomTag.Rows, "Rows", 1, 4096, sb);
                allPass &= AqCheckNumericRange(ds, DicomTag.Columns, "Columns", 1, 4096, sb);

                // 4) UID 형식 간단 검증 (숫자와 . 로만 구성)
                allPass &= AqCheckUidFormat(ds, DicomTag.StudyInstanceUID, "Study Instance UID", sb);
                allPass &= AqCheckUidFormat(ds, DicomTag.SeriesInstanceUID, "Series Instance UID", sb);
                allPass &= AqCheckUidFormat(ds, DicomTag.SOPInstanceUID, "SOP Instance UID", sb);

                string header = allPass ? "[AQ 결과] PASS" : "[AQ 결과] 경고/실패 있음";
                AddLog("AQ", $"{header} - {Path.GetFileName(path)}");

                // 세부 결과도 로그에 남기기
                foreach (var line in sb.ToString()
                                       .Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries))
                {
                    AddLog("AQ", line);
                }

                MessageBox.Show(
                    sb.ToString(),
                    header,
                    MessageBoxButton.OK,
                    allPass ? MessageBoxImage.Information : MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                AddLog("ERROR", $"AQ 실행 중 예외: {ex.Message}");
                MessageBox.Show(ex.ToString(), "AQ 오류");
            }
        }

        // ==========================
        // 공통: 폴더/스크롤/표시 로직
        // ==========================

        // 폴더 선택
        private string PickDicomFolder()
        {
            using (var dlg = new Forms.FolderBrowserDialog())
            {
                dlg.Description = "DICOM(.dcm) 파일이 있는 폴더를 선택하세요.";
                dlg.ShowNewFolderButton = false;

                var result = dlg.ShowDialog();
                if (result == Forms.DialogResult.OK && Directory.Exists(dlg.SelectedPath))
                    return dlg.SelectedPath;
            }
            return null;
        }

        // 폴더에서 .dcm 파일목록 로드
        private void LoadFolder(string folder)
        {
            _files = Directory.GetFiles(folder, "*.dcm")
                              .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                              .ToArray();
        }

        // 마우스 휠로 이전/다음 슬라이스
        private void Window_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (_files.Length == 0) return;

            if (e.Delta > 0) _index--;
            else _index++;

            if (_index < 0) _index = 0;
            if (_index >= _files.Length) _index = _files.Length - 1;

            // 스크롤바도 동기화
            _suppressScrollEvent = true;
            SliceScrollBar.Value = _index;
            _suppressScrollEvent = false;

            ShowDicom(_files[_index]);
        }

        // 스크롤바 드래그로 슬라이스 점프
        private void SliceScrollBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_files.Length == 0) return;
            if (_suppressScrollEvent) return;

            int newIndex = (int)Math.Round(SliceScrollBar.Value);
            if (newIndex < 0) newIndex = 0;
            if (newIndex >= _files.Length) newIndex = _files.Length - 1;

            _index = newIndex;
            ShowDicom(_files[_index]);
        }

        // 이미지 표시 + 태그 업데이트
        private void ShowDicom(string path)
        {
            try
            {
                // 먼저 DICOM 파일 열기
                var dicomFile = DicomFile.Open(path);
                var ds = dicomFile.Dataset;

                // Pixel Data(영상)가 없는 경우 : 태그만 보여주고 경고
                if (!ds.Contains(DicomTag.PixelData))
                {
                    DicomImageView.Source = null;   // 이미지 지우기
                    UpdateTags(ds);                // 태그는 그대로 표시

                    if (_files != null && _files.Length > 0)
                        Title = $"{System.IO.Path.GetFileName(path)} ({_index + 1}/{_files.Length}) - Pixel Data 없음";
                    else
                        Title = System.IO.Path.GetFileName(path) + " (Pixel Data 없음)";

                    // 같은 파일에 대해서는 한 번만 경고 표시
                    if (!string.Equals(_lastPixelDataWarningFile, path, StringComparison.OrdinalIgnoreCase))
                    {
                        _lastPixelDataWarningFile = path;

                        MessageBox.Show(
                            "이 DICOM 파일에는 Pixel Data(영상)가 없습니다.\n" +
                            "RT Plan / Report 등의 비영상 객체일 수 있습니다.",
                            "영상 없음",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
                    }

                    return;
                }

                // 여기부터는 Pixel Data가 있는 ‘영상’인 경우
                var dicomImage = new DicomImage(ds);
                var rendered = dicomImage.RenderImage(0);

                using (System.Drawing.Bitmap bmp = rendered.As<System.Drawing.Bitmap>())
                {
                    IntPtr hBitmap = bmp.GetHbitmap();
                    try
                    {
                        BitmapSource bs = Imaging.CreateBitmapSourceFromHBitmap(
                            hBitmap,
                            IntPtr.Zero,
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions());

                        bs.Freeze();
                        DicomImageView.Source = bs;
                    }
                    finally
                    {
                        DeleteObject(hBitmap);
                    }
                }

                // 태그 로드 & ListView 갱신
                UpdateTags(ds);

                // 타이틀 표시
                if (_files != null && _files.Length > 0)
                    Title = $"{System.IO.Path.GetFileName(path)} ({_index + 1}/{_files.Length})";
                else
                    Title = System.IO.Path.GetFileName(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "DICOM 표시 실패");
            }
        }

        // DICOM 태그를 Key/Value 형태로 ListView에 표시
        private void UpdateTags(DicomDataset ds)
        {
            TagItems.Clear();
            if (ds == null) return;

            foreach (var item in ds)
            {
                string key = item.Tag.DictionaryEntry != null
                    ? item.Tag.DictionaryEntry.Name
                    : item.Tag.ToString();

                // 값이 너무 길면 UI가 난장판 되니까 길이 제한
                string value;
                try
                {
                    value = ds.GetValueOrDefault(item.Tag, 0, "").ToString();
                }
                catch
                {
                    // 값이 배열이거나 복잡해서 GetValueOrDefault가 안 먹는 경우
                    value = item.ToString();
                }

                if (value == null) value = "";

                const int maxLen = 120;
                if (value.Length > maxLen)
                    value = value.Substring(0, maxLen) + " ...";

                TagItems.Add(new TagItem
                {
                    Key = key,
                    Value = value
                });
            }
        }

        // ======================
        // AQ용 헬퍼 메서드들
        // ======================

        // 1) 비식별화로 "삭제" 되어 있어야 하는 태그 검사
        private bool AqCheckRemovedTag(DicomDataset ds, DicomTag tag, string displayName, StringBuilder sb)
        {
            if (ds.Contains(tag))
            {
                string msg = $"[FAIL] {displayName} 태그가 아직 남아 있습니다. ({tag})";
                sb.AppendLine(msg);
                AddLog("AQ", msg);
                return false;
            }
            else
            {
                string msg = $"[OK] {displayName} 태그가 제거되었습니다.";
                sb.AppendLine(msg);
                AddLog("AQ", msg);
                return true;
            }
        }

        // 2) 반드시 존재해야 하는 문자열 태그 검사
        private bool AqCheckRequiredTagString(DicomDataset ds, DicomTag tag, string displayName, StringBuilder sb)
        {
            if (!ds.TryGetSingleValue(tag, out string value) || string.IsNullOrWhiteSpace(value))
            {
                string msg = $"[FAIL] {displayName} 태그가 비어있거나 없습니다. ({tag})";
                sb.AppendLine(msg);
                AddLog("AQ", msg);
                return false;
            }
            else
            {
                string msg = $"[OK] {displayName} = \"{value}\"";
                sb.AppendLine(msg);
                AddLog("AQ", msg);
                return true;
            }
        }

        // 3) 숫자 태그 범위 검사 (값이 있을 때만 체크, 없으면 패스)
        private bool AqCheckNumericRange(DicomDataset ds, DicomTag tag, string displayName,
                                         double min, double max, StringBuilder sb)
        {
            if (!ds.TryGetSingleValue(tag, out double value))
            {
                string msgSkip = $"[SKIP] {displayName} 값이 없어 범위 검사는 생략합니다. ({tag})";
                sb.AppendLine(msgSkip);
                AddLog("AQ", msgSkip);
                return true; // 값이 없다고 실패로 보지는 않음
            }

            if (value < min || value > max)
            {
                string msg = $"[WARN] {displayName} 값이 예상 범위를 벗어났습니다. ({value}, 기준: {min}~{max})";
                sb.AppendLine(msg);
                AddLog("AQ", msg);
                return false;
            }
            else
            {
                string msg = $"[OK] {displayName} = {value} (기준: {min}~{max})";
                sb.AppendLine(msg);
                AddLog("AQ", msg);
                return true;
            }
        }

        // 4) UID 문자열 형식 검사(숫자와 '.'만, 앞뒤에 '.' 없음)
        private bool AqCheckUidFormat(DicomDataset ds, DicomTag tag, string displayName, StringBuilder sb)
        {
            if (!ds.TryGetSingleValue(tag, out string uid) || string.IsNullOrWhiteSpace(uid))
            {
                string msg = $"[FAIL] {displayName} UID가 비어있습니다. ({tag})";
                sb.AppendLine(msg);
                AddLog("AQ", msg);
                return false;
            }

            bool isValid = uid.All(c => (c >= '0' && c <= '9') || c == '.') &&
                           !uid.StartsWith(".") &&
                           !uid.EndsWith("..") &&
                           !uid.Contains("..");

            if (!isValid)
            {
                string msg = $"[WARN] {displayName} UID 형식이 이상할 수 있습니다. ({uid})";
                sb.AppendLine(msg);
                AddLog("AQ", msg);
                return false;
            }
            else
            {
                string msg = $"[OK] {displayName} UID 형식 정상.";
                sb.AppendLine(msg);
                AddLog("AQ", msg);
                return true;
            }
        }

        // GDI 자원 정리
        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);
    }
}
