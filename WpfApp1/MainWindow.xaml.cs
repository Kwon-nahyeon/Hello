using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;

using Microsoft.Win32;

using Dicom;
using Dicom.Imaging;
using Dicom.IO.Buffer;

// 폴더 선택창
using Forms = System.Windows.Forms;

namespace WpfApp1
{
    // =======================
    // DICOM 비식별화 규칙 모음
    // =======================
    public static class DicomAnonymizer
    {
        public static void AnonymizeDataset(DicomDataset ds)
        {
            if (ds == null) return;

            // 삭제
            RemoveIfExists(ds, DicomTag.PatientName);
            RemoveIfExists(ds, DicomTag.PatientBirthDate);

            // PatientID는 "환자 단위 추적"을 위해
            // 삭제 대신 가명화를 택할 수 있음 (현재는 건드리지 않음)

            // 기타 환자 ID 범위 삭제
            RemovePatientOtherIDs(ds);

            // 날짜/시간 삭제
            RemoveIfExists(ds, DicomTag.StudyDate);
            RemoveIfExists(ds, DicomTag.StudyTime);
            RemoveIfExists(ds, DicomTag.SeriesDate);
            RemoveIfExists(ds, DicomTag.AcquisitionDate);

            // 기관/의료진 정보 삭제
            RemoveIfExists(ds, DicomTag.EthnicGroup);
            RemoveIfExists(ds, DicomTag.Occupation);
            RemoveIfExists(ds, DicomTag.InstitutionName);
            RemoveIfExists(ds, DicomTag.InstitutionAddress);
            RemoveIfExists(ds, DicomTag.ReferringPhysicianName);
            RemoveIfExists(ds, DicomTag.PerformingPhysicianName);
            RemoveIfExists(ds, DicomTag.OperatorsName);

            // UID 재생성
            RegenerateUid(ds, DicomTag.StudyInstanceUID);
            RegenerateUid(ds, DicomTag.SeriesInstanceUID);
            RegenerateUid(ds, DicomTag.SOPInstanceUID);
        }

        private static void RemoveIfExists(DicomDataset ds, DicomTag tag)
        {
            if (ds.Contains(tag)) ds.Remove(tag);
        }

        private static void RemovePatientOtherIDs(DicomDataset ds)
        {
            var toRemove = ds
                .Where(x => x.Tag.Group == 0x0010 &&
                            x.Tag.Element >= 0x1000 &&
                            x.Tag.Element <= 0x109C)
                .Select(x => x.Tag)
                .Distinct()
                .ToList();

            foreach (var tag in toRemove)
                ds.Remove(tag);
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
        public ObservableCollection<TagItem> TagItems { get; } = new ObservableCollection<TagItem>();

        // ====== 파일/시리즈 상태 ======
        private string[] _files = Array.Empty<string>();
        private int _axialIndex = 0;
        private bool _suppressScrollEvent = false;

        
        private static int Clamp(int v, int min, int max)
        {
            if (v < min) return min;
            if (v > max) return max;
            return v;
        }

        // Pixel Data 없는 파일 반복 팝업 방지용 (필요시 사용)
        private string _lastNoPixelWarnedPath = null;

        // ====== 볼륨 데이터(CT) ======
        private int _rows = 0;
        private int _cols = 0;
        private int _slices = 0;

        // slice마다 rows*cols 크기의 raw 값(int)
        private int[][] _volume = null;

        // Rescale (HU)
        private double _slope = 1.0;
        private double _intercept = 0.0;

        // Orthogonal plane 위치
        private int _sagittalX = 0; // 0 ~ cols-1
        private int _coronalY = 0;  // 0 ~ rows-1

        // 현재 표시 중인 파일 경로
        private string _currentPath = null;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            Loaded += MainWindow_Loaded;
        }

        // ======================
        // UI 유틸: 로그
        // ======================
        private void AddLog(string level, string message)
        {
            if (LogList == null) return;
            string line = $"[{DateTime.Now:HH:mm:ss}] [{level}] {message}";
            LogList.Items.Add(line);
            LogList.ScrollIntoView(line);
        }

        // ======================
        // 시작 시: 폴더 선택 (기존 흐름 유지)
        // ======================
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            string folder = PickDicomFolder();
            if (string.IsNullOrEmpty(folder))
            {
                Close();
                return;
            }

            LoadSeriesFromFolder(folder, preferredFilePath: null);
        }

        // ======================
        // 상단 버튼: 파일/폴더 열기
        // ======================
        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "DICOM 파일 선택",
                Filter = "DICOM (*.dcm)|*.dcm|All files (*.*)|*.*",
                Multiselect = false
            };

            if (dlg.ShowDialog() != true) return;

            string filePath = dlg.FileName;
            if (!File.Exists(filePath)) return;

            string folder = Path.GetDirectoryName(filePath);
            LoadSeriesFromFolder(folder, preferredFilePath: filePath);
        }

        private void OpenFolder_Click(object sender, RoutedEventArgs e)
        {
            string folder = PickDicomFolder();
            if (string.IsNullOrEmpty(folder)) return;

            LoadSeriesFromFolder(folder, preferredFilePath: null);
        }

        // ======================
        // 비식별 저장: 현재 파일 전체 저장
        // ======================
        private void Anonymize_Click(object sender, RoutedEventArgs e)
        {
            if (_files == null || _files.Length == 0)
            {
                MessageBox.Show("먼저 폴더를 열어 DICOM 파일들을 로드한 뒤 사용하세요.", "알림");
                return;
            }

            try
            {
                // 저장 폴더: 현재 시리즈 폴더 아래 Anonymized
                string baseDir = Path.GetDirectoryName(_files[0]);
                string outDir = Path.Combine(baseDir, "Anonymized");
                Directory.CreateDirectory(outDir);

                int success = 0;
                int fail = 0;

                AddLog("INFO", $"전체 비식별 저장 시작: 총 {_files.Length}개");

                foreach (var inputPath in _files)
                {
                    try
                    {
                        var dicomFile = DicomFile.Open(inputPath);

                        // 비식별화 규칙 적용
                        DicomAnonymizer.AnonymizeDataset(dicomFile.Dataset);

                        // 파일명 뒤에 _anon 붙여서 저장
                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(inputPath);
                        string ext = Path.GetExtension(inputPath);
                        string outputPath = Path.Combine(outDir, fileNameWithoutExt + "_anon" + ext);

                        dicomFile.Save(outputPath);
                        success++;
                    }
                    catch (Exception exOne)
                    {
                        fail++;
                        AddLog("WARN", $"비식별 저장 실패: {Path.GetFileName(inputPath)} / {exOne.Message}");
                    }
                }

                AddLog("INFO", $"전체 비식별 저장 완료: 성공 {success} / 실패 {fail} / 폴더: {outDir}");

                MessageBox.Show(
                    $"비식별 저장 완료!\n\n성공: {success}\n실패: {fail}\n\n저장 폴더:\n{outDir}",
                    "비식별 저장 완료",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AddLog("ERROR", $"전체 비식별 저장 중 오류: {ex.Message}");
                MessageBox.Show(ex.ToString(), "비식별 저장 실패");
            }
        }


        // ======================
        // AQ
        // ======================
        private void RunAQ_Click(object sender, RoutedEventArgs e)
        {
            if (_files == null || _files.Length == 0 || string.IsNullOrEmpty(_currentPath))
            {
                MessageBox.Show("먼저 DICOM 파일을 열어주세요.", "AQ", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string path = _currentPath;

            try
            {
                var dicomFile = DicomFile.Open(path);
                var ds = dicomFile.Dataset;

                var sb = new StringBuilder();
                bool allPass = true;

                sb.AppendLine($"[파일] {Path.GetFileName(path)}");

                // 제거되어야 하는 태그 확인
                allPass &= AqCheckRemovedTag(ds, DicomTag.PatientName, "Patient's Name", sb);
                allPass &= AqCheckRemovedTag(ds, DicomTag.PatientBirthDate, "Patient's Birth Date", sb);
                allPass &= AqCheckRemovedTag(ds, DicomTag.ReferringPhysicianName, "Referring Physician's Name", sb);
                allPass &= AqCheckRemovedTag(ds, DicomTag.InstitutionName, "Institution Name", sb);

                // 필수/유지 태그
                allPass &= AqCheckRequiredTagString(ds, DicomTag.PatientSex, "Patient's Sex", sb);
                allPass &= AqCheckRequiredTagString(ds, DicomTag.Modality, "Modality", sb);
                allPass &= AqCheckRequiredTagString(ds, DicomTag.StudyInstanceUID, "Study Instance UID", sb);
                allPass &= AqCheckRequiredTagString(ds, DicomTag.SeriesInstanceUID, "Series Instance UID", sb);
                allPass &= AqCheckRequiredTagString(ds, DicomTag.SOPInstanceUID, "SOP Instance UID", sb);

                // 수치
                allPass &= AqCheckNumericRange(ds, DicomTag.KVP, "KVP", 50, 150, sb);
                allPass &= AqCheckNumericRange(ds, DicomTag.Rows, "Rows", 1, 4096, sb);
                allPass &= AqCheckNumericRange(ds, DicomTag.Columns, "Columns", 1, 4096, sb);

                // UID 형식
                allPass &= AqCheckUidFormat(ds, DicomTag.StudyInstanceUID, "Study Instance UID", sb);
                allPass &= AqCheckUidFormat(ds, DicomTag.SeriesInstanceUID, "Series Instance UID", sb);
                allPass &= AqCheckUidFormat(ds, DicomTag.SOPInstanceUID, "SOP Instance UID", sb);

                string header = allPass ? "[AQ 결과] PASS" : "[AQ 결과] 경고/실패 있음";
                AddLog("AQ", $"{header} - {Path.GetFileName(path)}");

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

        // ======================
        // 3면/WWWC 로직 핵심
        // ======================

        private void LoadSeriesFromFolder(string folder, string preferredFilePath)
        {
            try
            {
                var files = Directory.GetFiles(folder, "*.dcm")
                                     .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                                     .ToArray();

                if (files.Length == 0)
                {
                    MessageBox.Show("선택한 폴더에 .dcm 파일이 없습니다.", "알림");
                    return;
                }

                _files = files;

                // 시작 인덱스 결정
                if (!string.IsNullOrEmpty(preferredFilePath))
                {
                    int idx = Array.FindIndex(_files, f => string.Equals(f, preferredFilePath, StringComparison.OrdinalIgnoreCase));
                    _axialIndex = (idx >= 0) ? idx : (_files.Length / 2);
                }
                else
                {
                    _axialIndex = _files.Length / 2;
                }

                // 시리즈(볼륨) 로드
                BuildVolumeFromFiles(_files);

                // 스크롤바 범위
                SliceScrollBar.Minimum = 0;
                SliceScrollBar.Maximum = Math.Max(0, _slices - 1);

                _suppressScrollEvent = true;
                SliceScrollBar.Value = _axialIndex;
                _suppressScrollEvent = false;

                // 사지탈/코로날 범위
                SagittalSlider.Minimum = 0;
                SagittalSlider.Maximum = Math.Max(0, _cols - 1);
                CoronalSlider.Minimum = 0;
                CoronalSlider.Maximum = Math.Max(0, _rows - 1);

                _sagittalX = _cols / 2;
                _coronalY = _rows / 2;

                SagittalSlider.Value = _sagittalX;
                CoronalSlider.Value = _coronalY;

                // WC/WW 초기값(태그에서 읽거나 기본값)
                InitWindowFromDicom(_files[Clamp(_axialIndex, 0, _slices - 1)]);

                RenderAll();

                AddLog("INFO", $"폴더 열기: {folder} (총 {_slices}장)");
            }
            catch (Exception ex)
            {
                AddLog("ERROR", $"폴더 로드 실패: {ex.Message}");
                MessageBox.Show(ex.ToString(), "폴더 로드 실패");
            }
        }

        private void BuildVolumeFromFiles(string[] files)
        {
            // Pixel Data 없는 파일은 자동 스킵해서 이미지 슬라이스만 모음
            var imageFiles = files.Where(IsPixelDataPresent).ToArray();
            if (imageFiles.Length == 0)
                throw new Exception("Pixel Data(영상)가 있는 DICOM이 없습니다. (RT Plan/Report 등일 수 있음)");

            // 기준 파일로 rows/cols
            var first = DicomFile.Open(imageFiles[0]);
            var ds0 = first.Dataset;

            _rows = ds0.GetSingleValueOrDefault(DicomTag.Rows, (ushort)0);
            _cols = ds0.GetSingleValueOrDefault(DicomTag.Columns, (ushort)0);

            if (_rows <= 0 || _cols <= 0)
                throw new Exception("Rows/Columns 정보를 읽을 수 없습니다.");

            // Rescale
            _slope = ds0.GetSingleValueOrDefault(DicomTag.RescaleSlope, 1.0);
            _intercept = ds0.GetSingleValueOrDefault(DicomTag.RescaleIntercept, 0.0);

            // 실제 볼륨 파일 목록은 영상 있는 파일로 교체
            _files = imageFiles;
            _slices = _files.Length;

            // axial index 범위 보정
            _axialIndex = Clamp(_axialIndex, 0, _slices - 1);

            _volume = new int[_slices][];

            for (int i = 0; i < _slices; i++)
            {
                var f = DicomFile.Open(_files[i]);
                var ds = f.Dataset;

                // PixelData 읽기
                var pixelData = DicomPixelData.Create(ds, false);
                if (pixelData.NumberOfFrames < 1)
                    throw new Exception($"프레임이 없습니다: {Path.GetFileName(_files[i])}");

                IByteBuffer frame = pixelData.GetFrame(0);
                byte[] bytes = frame.Data;

                int count = _rows * _cols;
                int[] slice = new int[count];

                ushort bitsAllocated = pixelData.BitsAllocated;          // 8 or 16
                int pixelRep = (int)pixelData.PixelRepresentation;      

                if (bitsAllocated == 16)
                {
                    if (pixelRep == 1) // signed
                    {
                        short[] tmp = new short[count];
                        Buffer.BlockCopy(bytes, 0, tmp, 0, Math.Min(bytes.Length, count * 2));
                        for (int p = 0; p < count; p++) slice[p] = tmp[p];
                    }
                    else // unsigned
                    {
                        ushort[] tmp = new ushort[count];
                        Buffer.BlockCopy(bytes, 0, tmp, 0, Math.Min(bytes.Length, count * 2));
                        for (int p = 0; p < count; p++) slice[p] = tmp[p];
                    }
                }
                else if (bitsAllocated == 8)
                {
                    for (int p = 0; p < Math.Min(bytes.Length, count); p++)
                        slice[p] = bytes[p];
                }
                else
                {
                    throw new Exception($"지원하지 않는 BitsAllocated={bitsAllocated}");
                }

                _volume[i] = slice;
            }
        }

        private bool IsPixelDataPresent(string path)
        {
            try
            {
                var f = DicomFile.Open(path);
                var ds = f.Dataset;
                return ds.Contains(DicomTag.PixelData);
            }
            catch
            {
                return false;
            }
        }

        private void InitWindowFromDicom(string path)
        {
            try
            {
                var f = DicomFile.Open(path);
                var ds = f.Dataset;

                double wc = ds.GetSingleValueOrDefault(DicomTag.WindowCenter, 40.0);
                double ww = ds.GetSingleValueOrDefault(DicomTag.WindowWidth, 400.0);

                if (double.IsNaN(wc)) wc = 40;
                if (double.IsNaN(ww) || ww < 1) ww = 400;

                WindowCenterSlider.Value = wc;
                WindowWidthSlider.Value = ww;
            }
            catch
            {
                WindowCenterSlider.Value = 40;
                WindowWidthSlider.Value = 400;
            }
        }

        private void RenderAll()
        {
            if (_volume == null || _volume.Length == 0) return;

            // Axial
            AxialImage.Source = RenderAxial(_axialIndex);

            // Sagittal/Coronal
            SagittalImage.Source = RenderSagittal(_sagittalX);
            CoronalImage.Source = RenderCoronal(_coronalY);

            // 현재 파일/태그
            _currentPath = _files[Clamp(_axialIndex, 0, _files.Length - 1)];
            UpdateTags(_currentPath);

            Title = $"{Path.GetFileName(_currentPath)} ({_axialIndex + 1}/{_slices})";
        }

        private BitmapSource RenderAxial(int z)
        {
            z = Clamp(z, 0, _slices - 1);
            return RenderGray(_volume[z], _cols, _rows);
        }

        private BitmapSource RenderSagittal(int x)
        {
            x = Clamp(x, 0, _cols - 1);

            int width = _slices;
            int height = _rows;
            int[] img = new int[width * height];

            for (int y = 0; y < _rows; y++)
            {
                for (int z = 0; z < _slices; z++)
                {
                    int axialIndex = y * _cols + x;
                    img[y * width + z] = _volume[z][axialIndex];
                }
            }
            return RenderGray(img, width, height);
        }

        private BitmapSource RenderCoronal(int y)
        {
            y = Clamp(y, 0, _rows - 1);

            int width = _slices;
            int height = _cols;
            int[] img = new int[width * height];

            for (int x = 0; x < _cols; x++)
            {
                for (int z = 0; z < _slices; z++)
                {
                    int axialIndex = y * _cols + x;
                    img[x * width + z] = _volume[z][axialIndex];
                }
            }
            return RenderGray(img, width, height);
        }

        private BitmapSource RenderGray(int[] raw, int width, int height)
        {
            double wc = WindowCenterSlider.Value;
            double ww = WindowWidthSlider.Value;
            if (ww < 1) ww = 1;

            byte[] pixels = new byte[width * height];

            double wcAdj = wc - 0.5;
            double wwAdj = ww - 1.0;

            for (int i = 0; i < pixels.Length; i++)
            {
                double hu = raw[i] * _slope + _intercept;
                double n = (hu - wcAdj) / wwAdj + 0.5;
                if (n < 0) n = 0;
                if (n > 1) n = 1;
                pixels[i] = (byte)(n * 255.0);
            }

            var wb = new WriteableBitmap(width, height, 96, 96, System.Windows.Media.PixelFormats.Gray8, null);
            wb.WritePixels(new Int32Rect(0, 0, width, height), pixels, width, 0);
            wb.Freeze();
            return wb;
        }

        // ======================
        // 이벤트: 슬라이스 / WWWC / 3면 위치
        // ======================
        // ======================
        // 이벤트: 슬라이스 / WWWC / 3면 위치
        // ======================
        private void SliceScrollBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_volume == null) return;
            if (_suppressScrollEvent) return;

            int newIndex = (int)Math.Round(SliceScrollBar.Value);
            newIndex = Clamp(newIndex, 0, _slices - 1);   

            _axialIndex = newIndex;
            RenderAll();
        }

        private void Axial_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (_volume == null) return;

            _axialIndex += (e.Delta > 0) ? -1 : 1;
            _axialIndex = Clamp(_axialIndex, 0, _slices - 1);  

            _suppressScrollEvent = true;
            SliceScrollBar.Value = _axialIndex;
            _suppressScrollEvent = false;

            RenderAll();
        }


        private void WindowSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_volume == null) return;
            RenderAll();
        }

        private void OrthoSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_volume == null) return;

            _sagittalX = (int)Math.Round(SagittalSlider.Value);
            _coronalY = (int)Math.Round(CoronalSlider.Value);

            // axial은 유지하고 사지탈/코로날만 다시 그려도 되지만 일단 단순하게 전체 렌더
            RenderAll();
        }

        // ======================
        // 태그 표시
        // ======================
        private void UpdateTags(string path)
        {
            TagItems.Clear();

            try
            {
                var dicomFile = DicomFile.Open(path);
                var ds = dicomFile.Dataset;

                foreach (var item in ds)
                {
                    string key = item.Tag.DictionaryEntry != null
                        ? item.Tag.DictionaryEntry.Name
                        : item.Tag.ToString();

                    string value;
                    try
                    {
                        value = ds.GetValueOrDefault(item.Tag, 0, "").ToString();
                    }
                    catch
                    {
                        value = item.ToString();
                    }

                    if (value == null) value = "";

                    const int maxLen = 160;
                    if (value.Length > maxLen)
                        value = value.Substring(0, maxLen) + " ...";

                    TagItems.Add(new TagItem { Key = key, Value = value });
                }
            }
            catch (Exception ex)
            {
                AddLog("WARN", $"태그 로드 실패: {ex.Message}");
            }
        }

        // ======================
        // 폴더 선택
        // ======================
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

        // ======================
        // AQ Helper
        // ======================
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

        private bool AqCheckNumericRange(DicomDataset ds, DicomTag tag, string displayName,
                                         double min, double max, StringBuilder sb)
        {
            if (!ds.TryGetSingleValue(tag, out double value))
            {
                string msgSkip = $"[SKIP] {displayName} 값이 없어 범위 검사는 생략합니다. ({tag})";
                sb.AppendLine(msgSkip);
                AddLog("AQ", msgSkip);
                return true;
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
                           !uid.Contains("..") &&
                           !uid.EndsWith(".");

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

        // ======================
        // GDI 자원 정리 (혹시 기존 Bitmap 변환 쓸 때 필요)
        // 지금은 WriteableBitmap으로 렌더해서 사실상 필요 없지만 혹시 몰라 남겨둠
        // ======================
        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);
    }
}
