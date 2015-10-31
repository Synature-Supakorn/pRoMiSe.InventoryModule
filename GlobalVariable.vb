Imports pRoMiSe.DBHelper
Imports System.Data.SqlClient

Public Class GlobalVariable

    Public DocDBUtil As CDBUtil
    Public DocConn As SqlConnection
    Public DefaultDocShopID As Integer = 1
    Public DocLangID As Integer = 2
    Public StaffID As Integer = 0
    Public DocYearSettingType As Integer = 1
    Public DefaultShopVAT As Decimal = 7
    Public InvariantCulture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.InvariantCulture
    Public DigitForRoundingDecimal As Integer = 2
    Public CurrencySymbol As String
    Public CurrencyCode As String
    Public CurrencyName As String
    Public CurrencyFormat As String
    Public DateFormat As String
    Public TimeFormat As String
    Public QtyFormat As String
    Public ShortDate As String
    Public ShortDateTime As String
    Public MaterialQtyFormat As String
    Public NumericFormat As String
    Public FullDateFormat As String
    Public FullDateTimeFormat As String
    Public AccountingFormat As String

    Public SoftwareVersion As String = "Build : 2015-10-16 V.1.0.27"
    '================== ORDERINVENTORY ===================='
    Public Const ORDERINVENTORY_BYNAME As Integer = 0
    Public Const ORDERINVENTORY_BYCODEANDNAME As Integer = 1
    Public Const ORDERINVENTORY_BYID As Integer = 2

    '================== Document Type ===================='
    Public Const DOCUMENTTYPE_PO As Integer = 1
    Public Const DOCUMENTTYPE_ROPO As Integer = 2
    Public Const DOCUMENTTYPE_TRANSFER As Integer = 3
    Public Const DOCUMENTTYPE_TRANSFERNOINVENTORY As Integer = 6
    Public Const DOCUMENTTYPE_MONTHLYSTOCK As Integer = 7
    Public Const DOCUMENTTYPE_RECEIPT As Integer = 8
    Public Const DOCUMENTTYPE_TRANSFERSTOCK As Integer = 10
    Public Const DOCUMENTTYPE_ROTRANSFER As Integer = 25
    Public Const DOCUMENTTYPE_REQUEST As Integer = 17
    Public Const DOCUMENTTYPE_DAILYSTOCK_ADD As Integer = 18
    Public Const DOCUMENTTYPE_DAILYSTOCK_REDUCE As Integer = 19
    Public Const DOCUMENTTYPE_SALE As Integer = 20
    Public Const DOCUMENTTYPE_VOID As Integer = 21
    Public Const DOCUMENTTYPE_MONTHLYSTOCK_ADD As Integer = 22
    Public Const DOCUMENTTYPE_MONTHLYSTOCK_REDUCE As Integer = 23
    Public Const DOCUMENTTYPE_DAILYSTOCK As Integer = 24
    Public Const DOCUMENTTYPE_WEEKLYSTOCK As Integer = 30
    Public Const DOCUMENTTYPE_WEEKLYSTOCK_ADD As Integer = 31
    Public Const DOCUMENTTYPE_WEEKLYSTOCK_REDUCE As Integer = 32
    Public Const DOCUMENTTYPE_DIRECTRO As Integer = 39
    Public Const DOCUMENTTYPE_DIRECTROPTT As Integer = 40
    Public Const DOCUMENTTYPE_DIRECTROPTTNONOIL As Integer = 41
    Public Const DOCUMENTTYPE_TRANSFERWITHCOST As Integer = 46
    Public Const DOCUMENTTYPE_ROTRANSFERWITHCOST As Integer = 47
    Public Const DOCUMENTTYPE_ROCREATEPO As Integer = 56
    Public Const DOCUMENTTYPE_ADJUSTSTOCK As Integer = 57
    Public Const DOCUMENTTYPE_ADJUSTSTOCK_ADD As Integer = 58
    Public Const DOCUMENTTYPE_ADJUSTSTOCK_REDUCE As Integer = 59
    Public Const DOCUMENTTYPE_TRANSFERBAKERY_DORO As Integer = 1000
    Public Const DOCUMENTTYPE_REQUESTBAKERY_DORO As Integer = 1002

    '================ Document Status =================='
    Public Const DOCUMENTSTATUS_TEMP As Integer = 0
    Public Const DOCUMENTSTATUS_WORKING As Integer = 1
    Public Const DOCUMENTSTATUS_APPROVE As Integer = 2
    Public Const DOCUMENTSTATUS_REFERED As Integer = 3
    Public Const DOCUMENTSTATUS_FINISH As Integer = 4
    Public Const DOCUMENTSTATUS_CANCEL As Integer = 99
    Public Const DOCUMENTSTATUS_NODOCSELECT As Integer = -1

    Public Const TAXTYPE_NOVAT As Integer = 0
    Public Const TAXTYPE_INCLUDEVAT As Integer = 1
    Public Const TAXTYPE_EXCLUDEVAT As Integer = 2

    '==================== Error Msg ===================='
    Public Const MESSAGE_DATANOTFOUND As String = "ไม่พบข้อมูลในระบบ"
    Public Const MESSAGE_MATERIALNOTFOUND As String = "ไม่พบข้อมูลสินค้าในระบบ"
    Public Const MESSAGE_NOTENOUGHSTOCK As String = "สต๊อกสินค้าในระบบไม่เพียงพอ"
    Public Const MESSAGE_FORWARDSTOCKS As String = "ไม่สามารถทำการบันทึกข้อมูลเอกสารนี้ได้ เนื่องจากคลังสินค้าได้มีการยกยอดสต๊อกผ่านเดือนของวันที่ของเอกสารมาแล้ว"
    Public Const MESSAGE_INVALIDDATECOUNTSTOCK As String = "วันที่เอกสารนับสต๊อกไม่ถูกต้อง"
    Public Const MESSAGE_MATERIALBELOWZERO As String = "พบสินค้าบางรายการที่มีจำนวนติดลบในระบบ จึงไม่สามารถอนุมัติเอกสารนี้ได้"

End Class
