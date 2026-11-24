# CVD Report

뇌심혈관질환 위험성평가 보고서 자동작성 프로그램.

## Windows EXE 자동 빌드

1. 이 폴더의 파일을 GitHub 저장소에 업로드합니다.
2. GitHub 저장소의 **Actions** 탭에서 **Build Windows EXE**를 선택합니다.
3. **Run workflow**를 눌러 빌드를 시작합니다.
4. 완료된 실행에서 `CVD_Report_Windows_x64_OCR` artifact ZIP을 다운로드합니다.
5. ZIP 압축을 풀고 `CVD_Report.exe`를 실행합니다.

`Tesseract-OCR` 실행파일과 한국어(`kor`)·영어(`eng`) 학습 데이터는 Windows 빌드 과정에서 함께 포함됩니다. 별도 Tesseract 설치 없이 OCR을 사용할 수 있도록 구성했습니다.

## 주의

- 배포 ZIP은 `onedir` 방식입니다. `CVD_Report.exe`만 따로 옮기지 말고 압축을 푼 폴더 전체를 유지하세요.
- 개인 건강검진 자료, 생성된 보고서, 개인별 저장 JSON은 GitHub에 업로드하지 마세요.
- OCR 결과는 반드시 원본 자료와 대조·확인해야 합니다.
