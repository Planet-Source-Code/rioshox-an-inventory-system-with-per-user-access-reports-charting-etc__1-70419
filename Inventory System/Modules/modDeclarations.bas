Attribute VB_Name = "modDeclarations"
Option Explicit

Public Const SUB_MENU_HEIGHT        As Long = 860
Public Const DB_NAME                As String = "\RFM.mdb"
Public Const COMPANY_NAME           As String = "RFM"
Public Const SYSTEM_NAME            As String = "Sales and Inventory System"


Public Const TO_ADD                 As Long = 1000
Public Const TO_EDIT                As Long = 2000

Public Const ACTIVE_MENU_WIDTH  As Long = 390

Public ActiveFrame%, ActiveTop%

Public isFullForm As Boolean 'Tracks if Sub mEnu is to be show in full mode
Public isClicked As Boolean
Public ActiveForm As Form
Public IsThereLoaded As Boolean
Public Const TO_ACCESS_COUNT As Integer = 8

Public LineDiscount As Double

Public connRFM                 As New ADODB.Connection




Type UsersInfo
    UserName As String
    EmployeeID As Integer
    IsAdmin As Boolean
    Products As Boolean
    Customer As Boolean
    UserAccess As Boolean
    ApproveSO As Boolean
    ApproveShip As Boolean
    ReturnItem As Boolean
    ManualSO As Boolean
    AuthorizeDiscount As Boolean
    BackupData As Boolean
End Type

Public g_CurrentUser As UsersInfo

Public dMenu As Integer
