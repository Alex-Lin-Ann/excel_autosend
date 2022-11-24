# 這是一個以前在前公司做的一個Outlook自動發信專案，利用Excel 的特性執行以下步驟:
1. 將系統裡的提單號碼和Email資料直接貼上
2. 並自動產生附件等相關文件
3. 最後搜尋資料是否都建立完成
4. 並批次寄發通知客戶
5. 等待客戶回傳相關資料，方可進行最後的作業流程

![image](https://user-images.githubusercontent.com/37874182/203689623-63d6f544-46cb-4265-9e5c-82261513cd04.png)

# 功能介紹
* A欄 - 貼上提單號碼
* B欄 - 貼上要寄發的Email
* C欄 - 輸入要在標題後方額外通知客戶的訊息
* D欄 - 輸入要在內容裡通知客戶的訊息
* E欄 - 雙擊兩下夾帶其他附件(可留空白)
* F欄 - 點擊右方按鈕[Search Data] 搜尋短缺的附件，如(提單與發票.pdf，自動產生個案委任書.doc，自動產生身分證正反面黏貼處.doc，智慧型手機 下在APP委任更方便.pdf)
* G欄 - 點擊右方按鈕[Send Mail]進行寄發的動作，寄發前會跳出MsgBox通知是否寄出並強制檢查F欄狀態為 (OK綠底) 字樣，並在背景啟動Outlook寄發郵件
* L1欄 - 輸入要發送Email需要CC的相關人
* L5欄 - 會出現在Outlook簽名檔上的名字，如(Alex Lin)
* L6欄 - 會出現在Outlook簽名檔上的職稱，如(Manager)
* L7欄 - 會出現在Outlook簽名檔上的公司電話，如(886-2123-xxxx)
* L8欄 - 會出現在Outlook簽名檔上的公司傳真，如(886-2123-xxxx)
* L9欄 - 會出現在Outlook簽名檔上的公司Email，如(hello1234@gmail.com)
* 按鈕 [Send Mail] - 啟動Outlook 並於背景寄發Email
* 按鈕 [Search Data] - 搜尋應夾帶的檔案，如(提單與發票.pdf，自動產生個案委任書.doc，自動產生身分證正反面黏貼處.doc，智慧型手機 下在APP委任更方便.pdf)
* 按鈕 [Produce POA] - A欄要貼上提單號碼後，才能夠點擊該按鈕，並批次完成製作個案委任書.doc 檔案
* 按鈕 [Produce PasteID] - A欄要貼上提單號碼後，才能夠點擊該按鈕，並批次完成身分證正反面黏貼處.doc 檔案
* 按鈕 [Clear Data] - 刪除於C槽中自動產生的個案委任書.doc和身分證正反面黏貼處.doc
* 按鈕 [Clear All] - 清空A2:G999之儲存格
* 按鈕 [Setup] - 第一次使用的人，自動在C槽產生Auto_POA和Auto_PasteID的資料夾

# Step by Step 使用教學
1. 請在A欄貼上提單號碼
2. 請在B欄貼上應寄發的Email
3. 請點選按鈕[Produce POA]自動產生個案委任書.doc
4. 請點選按鈕[Produce PasteID]自動產生身分證正反面黏貼處.doc
5. 請點選按鈕[Search Data]搜尋自動產生的檔案
6. 請點選按鈕[Send Mail]批次寄送Email

# The Above
因為已經是5年前做的簡單小工具，登入密碼也已經遺忘了，也沒有再進行維護，上述說明簡單的介紹大致上功能
