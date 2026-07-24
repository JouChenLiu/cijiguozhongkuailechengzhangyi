# cijiguozhongkuailechengzhangyi
慈濟暑期國中快樂成長營

[國中營現場即時管理系統](https://jouchenliu.github.io/cijiguozhongkuailechengzhangyi/)

## 系統使用者權限設定
> In Firebase: get User ID in the Authentication; set permissions in the Realtime Database's users list.

userRole:
* admin: 有最大權限的管理員。可編輯、查看所有資料(包含營隊詳細數據分析、聯繫訊息等)、設定開放年度資料。<br>
  `目前擁有此權限的帳號: 國中營Google帳號、Anna的帳密`
* renzhen: 有第二大權限的使用者，用於管理學員的受訪同意書(人真組)。<br>
  `目前尚無設定帳號擁有此權限`
* reader: 無須登入帳號的一般使用者。只能查看一般名單資料(含簡易的報到人數統計)，無法做任何編輯。可使用右下角的訊息功能，聯繫管理員。<br>
  `一般使用者僅能查看admin已設定開放年度的名單資料`

## 管理對象與(後端)資料欄位
* currentyear/campers: 學員名單(含隊輔)
  * 報到、發票、同意書、身分、繳費狀況、組別、姓名、性別、待確認事項、報到備註、備註
* currentyear/staff: 慈青工作人員
  * 報到、職務、身分別、組別、姓名、生理性別、全程參與、備註

## 系統功能

一般使用者&管理員
* 查詢/篩選 名單
  * 選擇年度
  1. 選擇管理對象 (學員名單(含隊輔) or 慈青工作人員)
  2. 篩選組別 (全部組別 or 尚未報到 or ...)
  3. 搜尋姓名
* 查看 簡易版報到人數統計
  * 學員 -> 點選"全體學員總計 (不含隊輔)"
  * 工作人員 -> 點選"全體工作團隊總計"
 
管理員
* 編輯 名單部分資訊(欄位)
  * 學員: 報到、發票、同意書、報到備註、待確認事項、✍️ 現場備註
  * 工作人員: 報到、⏱️ 全程參與、✍️ 現場備註
* 編輯 開放年度設定 (有開啟權限的年度，一般使用者才看得到該年度的資料)
* 查看 營隊詳細數據分析
* 查看 聯繫訊息

## 系統架構
* 前端: index.html
* 後端: Firebase
* 資料: 從人票總表 > 透過upload_excel.py > 匯入Firebase
* 架設平台: GitHub
* 版本控制: Git

## 開發者
Anna and Gemini Pro

(註: Anna有保護大家的個資 沒有洩漏給AI)
