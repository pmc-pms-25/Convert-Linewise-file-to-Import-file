# Convert MTO (Piping / Structural / EIT)

Tool chuyển file MTO do DE nhập sang template nạp hệ thống PMS / Workfront.

| Bộ môn | Đầu vào | Kết quả |
|---|---|---|
| Piping (PIP) | File MTO bất kỳ định dạng (AVEVA linewise, Import CSV…), map cột theo từ điển | Template Import 36 cột (sheet Import của Working_Nguyen): FabricationType, UnitDiaInch theo NEW Typical, Quantity |
| Structural (STR) | Structural MTO Master | Template StructuralMTOMaster (sheet DATA), WorkGroup tra theo config |
| EIT | EIT MTO Master: Support / Tray-Ladder / Equipment / Tubing / Cable schedule | Layout mtoc_eit_mto_master, tự điền DisciplineType theo loại, WorkGroup suy theo loại vật tư |

File chỉ được xử lý trên máy (trong trình duyệt), không gửi đi đâu.

## Dùng tool

- `Convert MTO.exe`: bấm đúp, mở tool trong cửa sổ riêng bằng Microsoft Edge có sẵn; chạy được khi không có mạng.
- `Convert MTO.html`: cùng tool, mở bằng Edge / Chrome (cần mạng để tải thư viện SheetJS).
- `TuDien_Cot.xlsx`: từ điển cột Piping, nạp vào tab "PIP · Từ điển cột".
- Hướng dẫn chi tiết: `HUONG DAN.txt`.

## Cấu trúc source

```
Convert MTO.exe / Convert MTO.html   bản build (tạo bởi src/build.ps1)
src/
  convert_ui_template.html   giao diện + engine Piping (/*DATA*/, /*MASTER*/ là chỗ chèn khi build)
  master.js                  engine STR / EIT, bảng config & tra cứu
  data/typical_data.js       NEW Typical (UnitDiaInch) nhúng sẵn
  data/lookup_data.js        config WorkGroup, sys_project_list, sys_facility_list, template input STR/EIT
  exe/Launcher.cs            exe: giải nén trang vào %LOCALAPPDATA%\ConvertMTO rồi mở Edge --app
  exe/app.ico, exe/xlsx.full.min.js (SheetJS 0.18.5)
  build.ps1                  build HTML + exe
  tools/export_lookup.ps1    tạo lại data/lookup_data.js từ Excel (config, project, facility, template input)
  tools/export_typical.ps1   tạo lại data/typical_data.js từ sheet NEW Typical
  tools/test_page.ps1        self-test headless bằng Edge
  tools/selftest.js          hook self-test chèn vào trang test
```

## Build

Yêu cầu: Windows, PowerShell 5.1, .NET Framework 4 (csc.exe có sẵn), Microsoft Edge. Không cần Python / Node.

```powershell
powershell -ExecutionPolicy Bypass -File src\build.ps1
```

Tạo `Convert MTO.html`, `Convert MTO.exe` ở thư mục gốc; file trung gian ở `src\build\` (bản online cho claude.ai là `src\build\convert-import.html`).

Cập nhật dữ liệu nhúng sẵn (cần Excel; các script chỉ mở bản sao tạm của file nguồn):

```powershell
powershell -ExecutionPolicy Bypass -File src\tools\export_lookup.ps1   # config WorkGroup, project, facility, template input
powershell -ExecutionPolicy Bypass -File src\tools\export_typical.ps1  # NEW Typical
powershell -ExecutionPolicy Bypass -File src\build.ps1
```

Người dùng cũng cập nhật được config / project / facility / Typical ngay trong tool (tab "Bảng config & tra cứu", "PIP · Cài đặt & Typical") mà không cần build lại.

## Self-test

```powershell
.\src\tools\test_page.ps1 -File "C:\path\EIT MTO Master (Tubing).xlsx" -Name tub -Query '&mode=eit' -B64
.\src\tools\test_page.ps1 -Page app-test.html -File "C:\path\Structural MTO Master.xlsx" -Name str -Query '&mode=str'
.\src\tools\test_page.ps1 -File "C:\path\linewise.xlsx" -Name pip -Query '&sys=2&pc=PKMD2515'
```

`-B64` lưu file kết quả vào `src\build\out\`.

## Quy tắc convert chính

- WorkGroup: ProjectCode → ProjectType (sys_project_list); ProjectCode + FacilityCode → FacilityType (sys_facility_list); tra `plan_config_work_group` theo DisciplineCode + ProjectType + FacilityType + WorkGroupName. STR ghi mã 3 số (001), EIT ghi số trơn. Không tra được / tên trùng nhiều mã → để trống, tô hồng, ghi lý do ở cột `Remark convert`.
- EIT WorkGroupName: Support → Prefab Support (UNIVERSAL → Prefab Universal Support), Tray/Ladder → Tray/Ladder, Equipment → Electrical/Instrument Equipment, Tubing → Tubing, Cable → LV Cable (≤ 1 kV) / HV/MV Cable (ELE), Cable (INS).
- Location: tên deck đầy đủ → mã (Cellar Deck → CED, Main Deck → MAD…); không đổi được → để trống + remark.
- Chi tiết từng cột: xem bảng "Cách map cột" trong tool hoặc sheet `Map` của file kết quả.