Attribute VB_Name = "M35_Models"
'============================================================
' M35_Models : パース・判定に使うデータモデル（UDT/Enum）
'============================================================
Option Explicit

' 異動の種別（ざっくり分類）
Public Enum MovementKind
    mvUnknown = 0
    mvBorn = 1          ' 出生
    mvPurchaseIn = 2    ' 導入/転入/購入/入牧
    mvSaleOut = 3       ' 売却/転出/出荷
    mvDeath = 4         ' 死亡
    mvTransfer = 5      ' その他の移動（施設内/施設間など）
End Enum

' 異動1件ぶん
Public Type MovementRow
    No As String                ' 番号（サイトの連番など）
    Kind As MovementKind        ' 種別（上のEnum）
    DateValue As Date           ' 異動年月日
    Pref As String              ' 飼養施設都道府県
    City As String              ' 飼養施設市区町村
    Name As String              ' 氏名または名称（牧場名）
    RawKind As String           ' 生の「異動内容」文字列（保全）
End Type

' 個体プロフィール
Public Type CowProfile
    CowID As String                 ' 個体識別番号（10桁）
    BirthDate As Date               ' 出生の年月日
    Sex As String                   ' 雌雄の別
    MotherID As String              ' 母牛の個体識別番号
    Breed As String                 ' 種別
End Type

' 判定結果（個体ごと）
Public Type CowRecord
    Profile As CowProfile
    Movements() As MovementRow

    ' 判定フラグ
    BornAtFarm As Boolean
    AtStart As Boolean
    BornInTerm As Boolean
    PurchasedInTerm As Boolean
    SoldInTerm As Boolean
    DiedInTerm As Boolean
    AtEnd As Boolean

    ' 代表日付・相手先（帳票用に便利）
    PurchaseDate As Variant ' Date or Empty
    PurchaseFrom As String
    SaleDate As Variant
    SaleTo As String
    DeathDate As Variant
End Type

