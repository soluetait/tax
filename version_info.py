"""PyInstaller용 버전 정보 파일 생성."""
import PyInstaller.utils.win32.versioninfo as vi

version = (1, 3, 0, 0)

info = vi.VSVersionInfo(
    ffi=vi.FixedFileInfo(
        filevers=version,
        prodvers=version,
        mask=0x3F,
        flags=0x0,
        OS=0x40004,
        fileType=0x1,
        subtype=0x0,
    ),
    kids=[
        vi.StringFileInfo([
            vi.StringTable("041204B0", [  # 0412 = Korean, 04B0 = Unicode
                vi.StringStruct("CompanyName", "(주)다산솔루에타"),
                vi.StringStruct("FileDescription", "AI 세금계산서"),
                vi.StringStruct("FileVersion", "1.3.0"),
                vi.StringStruct("InternalName", "AI세금계산서"),
                vi.StringStruct("LegalCopyright", "(c) 2026 다산솔루에타"),
                vi.StringStruct("OriginalFilename", "AI세금계산서.exe"),
                vi.StringStruct("ProductName", "AI 세금계산서"),
                vi.StringStruct("ProductVersion", "1.3.0"),
            ])
        ]),
        vi.VarFileInfo([vi.VarStruct("Translation", [0x0412, 0x04B0])])
    ]
)

if __name__ == "__main__":
    with open("version_info.txt", "w", encoding="utf-8") as f:
        f.write(str(info))
    print("version_info.txt 생성 완료")
