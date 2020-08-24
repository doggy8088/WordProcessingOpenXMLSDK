# 自動調整 Word 文件說明

- 文件範本

    ```sh
    程式異動申請清單.tpl.docx
    ```

- 執行方式

    直接讀取 `Program.cs` 程式碼中的 `text` 字串內容 (測試用途)

    ```sh
    dotnet run
    ```

    直接讀取 Clipboard 剪貼簿中的文字 (Windows-only)

    ```sh
    dotnet run -c Release
    ```

    > 請從 TortoiseGit 複製完整檔案更新資訊
    >
    > ![從 TortoiseGit 複製完整檔案更新資訊](https://user-images.githubusercontent.com/88981/91070939-17e49680-e66a-11ea-9998-292be16a49cf.png)

- 發行專案

    ```sh
    dotnet publish -c Release -r win10-x64 "-p:PublishSingleFile=true" "-p:PublishTrimmed=true" -o dist\
    ```

    主程式將會輸出到 `dist\WordProcessingOpenXMLSDK.exe`

- 使用方式說明

    1. 將 `dist\WordProcessingOpenXMLSDK.exe` 執行檔與 `程式異動申請清單.tpl.docx` 文件範本放在同一個資料夾下

    2. 使用 TortoiseGit 比對兩個不同的版本，並且比對差異檔案，然後複製所有資訊到剪貼簿。

    3. 直接執行 `WordProcessingOpenXMLSDK.exe` 就可以自動產生 `程式異動申請清單.docx` 文件！

- 使用教學影片

    <https://youtu.be/YM9NfcDsGs0>

## 相關連結

- [Word processing (Open XML SDK)](https://docs.microsoft.com/zh-tw/office/open-xml/word-processing)
- [Insert a table into a word processing document (Open XML SDK)](https://docs.microsoft.com/zh-tw/office/open-xml/how-to-insert-a-table-into-a-word-processing-document)
- [OpenXML replace text in all document](https://stackoverflow.com/a/19100156/910074)
- [TextCopy - A cross platform package to copy text to and from the clipboard.](https://github.com/CopyText/TextCopy/)