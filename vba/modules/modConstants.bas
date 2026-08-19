Attribute VB_Name = "modConstants"
Option Explicit

' ============================================================================
' modConstants - Frozen application contract constants (D-021)
' ============================================================================
' Every literal referenced by VBA that comes from the frozen contract lives
' here so a single place mirrors schema/workbook-contract.yaml (status: frozen,
' contract_version 1.0.0). Change control applies: see docs/workbook-contract.md.
' ============================================================================

' ------------------------------------------------------------------ contract
Public Const CONTRACT_VERSION As String = "1.0.0"
Public Const WORKBOOK_VERSION As String = "v0.1"

' ------------------------------------------------------------------ worksheets
Public Const WS_DASHBOARD As String = "Dashboard"
Public Const WS_SCAN As String = "Scan"
Public Const WS_RECEIVING As String = "Receiving"
Public Const WS_PRODUCTS As String = "Products"
Public Const WS_CONTAINERS As String = "Containers"
Public Const WS_TRANSACTIONS As String = "Transactions"
Public Const WS_SUPPLIERS As String = "Suppliers"
Public Const WS_LOCATIONS As String = "Locations"
Public Const WS_SETTINGS As String = "Settings"

' ------------------------------------------------------------------ tables
Public Const TBL_PRODUCTS As String = "tblProducts"
Public Const TBL_CONTAINERS As String = "tblContainers"
Public Const TBL_TRANSACTIONS As String = "tblTransactions"
Public Const TBL_SUPPLIERS As String = "tblSuppliers"
Public Const TBL_LOCATIONS As String = "tblLocations"
Public Const TBL_SETTINGS As String = "tblSettings"
Public Const TBL_STATUS_LIST As String = "tblStatusList"
Public Const TBL_TXN_TYPE_LIST As String = "tblTransactionTypeList"
Public Const TBL_EXPIRY_CLASS_LIST As String = "tblExpiryClassList"
Public Const TBL_SCAN_RESULTS As String = "tblScanResults"
Public Const TBL_RECEIVE_STAGING As String = "tblReceiveStaging"

' ------------------------------------------------------------------ tblProducts columns
Public Const COL_PRODUCT_ID As String = "ProductID"
Public Const COL_PRODUCT_NAME As String = "ProductName"
Public Const COL_PRODUCT_TYPE As String = "ProductType"
Public Const COL_CATEGORY As String = "Category"
Public Const COL_MANUFACTURER As String = "Manufacturer"
Public Const COL_MANUFACTURER_CATNO As String = "ManufacturerCatalogueNumber"
Public Const COL_CAS_NUMBER As String = "CASNumber"
Public Const COL_CONCENTRATION As String = "Concentration"
Public Const COL_GRADE As String = "Grade"
Public Const COL_STD_CONTAINER_DESC As String = "StandardContainerDescription"
Public Const COL_SUPPLIER_ID As String = "SupplierID"
Public Const COL_STORAGE_REQUIREMENTS As String = "StorageRequirements"
Public Const COL_HAZARD_CLASS As String = "HazardClassification"
Public Const COL_SDS_REF As String = "SDSReference"
Public Const COL_MIN_STOCK As String = "MinimumContainerStock"
Public Const COL_TARGET_STOCK As String = "TargetContainerStock"
Public Const COL_REORDER_QTY As String = "ReorderQuantity"
Public Const COL_ACTIVE As String = "Active"
Public Const COL_NOTES As String = "Notes"
Public Const COL_HELPER_AVAILABLE_STOCK As String = "HelperAvailableStock"
Public Const COL_HELPER_STOCK_CLASS As String = "HelperStockClass"

' ------------------------------------------------------------------ tblContainers columns
Public Const COL_CONTAINER_ID As String = "ContainerID"
Public Const COL_BARCODE As String = "Barcode"
Public Const COL_BATCH_LOT As String = "BatchLotNumber"
Public Const COL_EXPIRY_DATE As String = "ExpiryDate"
Public Const COL_RETEST_DATE As String = "RetestDate"
Public Const COL_DATE_RECEIVED As String = "DateReceived"
Public Const COL_STORAGE_LOCATION_ID As String = "StorageLocationID"
Public Const COL_STATUS As String = "Status"
Public Const COL_OPENED_DATE As String = "OpenedDate"
Public Const COL_DISPOSAL_DATE As String = "DisposalDate"
Public Const COL_DISPOSAL_REASON As String = "DisposalReason"
Public Const COL_HELPER_CONTAINER_NUM As String = "HelperContainerNum"
Public Const COL_HELPER_BARCODE_NUM As String = "HelperBarcodeNum"

' ------------------------------------------------------------------ tblTransactions columns
Public Const COL_TRANSACTION_ID As String = "TransactionID"
Public Const COL_TIMESTAMP As String = "Timestamp"
Public Const COL_OPERATOR As String = "Operator"
Public Const COL_TXN_TYPE As String = "TransactionType"
Public Const COL_PREVIOUS_STATUS As String = "PreviousStatus"
Public Const COL_NEW_STATUS As String = "NewStatus"
Public Const COL_PREVIOUS_LOCATION As String = "PreviousLocation"
Public Const COL_NEW_LOCATION As String = "NewLocation"
Public Const COL_REASON As String = "Reason"
Public Const COL_REFERENCE As String = "Reference"

' ------------------------------------------------------------------ named ranges
Public Const RNG_SCAN_INPUT As String = "rngScanInput"
Public Const RNG_SCAN_STATUS_MESSAGE As String = "rngScanStatusMessage"
Public Const RNG_SCAN_RESULT_CARD As String = "rngScanResultCard"
Public Const RNG_RECV_PRODUCT_ID As String = "rngReceiveProductID"
Public Const RNG_RECV_NEXT_CONTAINER_ID As String = "rngReceiveNextContainerID"
Public Const RNG_RECV_NEXT_BARCODE As String = "rngReceiveNextBarcode"
Public Const RNG_RECV_LOT As String = "rngReceiveLot"
Public Const RNG_RECV_EXPIRY As String = "rngReceiveExpiry"
Public Const RNG_RECV_RETEST As String = "rngReceiveRetest"
Public Const RNG_RECV_LOCATION As String = "rngReceiveLocation"
Public Const RNG_RECV_QUANTITY As String = "rngReceiveQuantity"
Public Const RNG_RECV_STATUS_MESSAGE As String = "rngReceiveStatusMessage"
Public Const NAME_DEFAULT_LOCATION As String = "DefaultLocationID"
Public Const NAME_DEFAULT_NEW_STATUS As String = "DefaultStatusNewContainers"
Public Const NAME_EXPIRY_30 As String = "ExpiryWarningDays30"
Public Const NAME_EXPIRY_60 As String = "ExpiryWarningDays60"
Public Const NAME_EXPIRY_90 As String = "ExpiryWarningDays90"

' ------------------------------------------------------------------ settings keys
Public Const SETTING_EXPIRY_30 As String = "ExpiryWarningDays30"
Public Const SETTING_EXPIRY_60 As String = "ExpiryWarningDays60"
Public Const SETTING_EXPIRY_90 As String = "ExpiryWarningDays90"
Public Const SETTING_DEFAULT_LOCATION As String = "DefaultLocationID"
Public Const SETTING_DEFAULT_NEW_STATUS As String = "DefaultStatusNewContainers"
Public Const SETTING_SCANNER_ENTER_SUFFIX As String = "ScannerEnterSuffix"
Public Const SETTING_WORKBOOK_VERSION As String = "WorkbookVersion"

' ------------------------------------------------------------------ statuses (frozen 6-value model, D-016)
Public Enum ContainerStatus
    stAvailable = 1
    stInUse = 2
    stExpired = 3
    stDamaged = 4
    stDisposed = 5
    stMissing = 6
End Enum

Public Const STATUS_AVAILABLE As String = "Available"
Public Const STATUS_IN_USE As String = "InUse"
Public Const STATUS_EXPIRED As String = "Expired"
Public Const STATUS_DAMAGED As String = "Damaged"
Public Const STATUS_DISPOSED As String = "Disposed"
Public Const STATUS_MISSING As String = "Missing"

' ------------------------------------------------------------------ transaction types (frozen 9-value model)
Public Enum TransactionType
    ttReceive = 1
    ttTakeOpen = 2
    ttReturn = 3
    ttTransfer = 4
    ttDispose = 5
    ttMarkExpired = 6
    ttMarkDamaged = 7
    ttMarkMissing = 8
    ttAdjustment = 9
End Enum

Public Const TXN_RECEIVE As String = "Receive"
Public Const TXN_TAKE_OPEN As String = "TakeOpen"
Public Const TXN_RETURN As String = "Return"
Public Const TXN_TRANSFER As String = "Transfer"
Public Const TXN_DISPOSE As String = "Dispose"
Public Const TXN_MARK_EXPIRED As String = "MarkExpired"
Public Const TXN_MARK_DAMAGED As String = "MarkDamaged"
Public Const TXN_MARK_MISSING As String = "MarkMissing"
Public Const TXN_ADJUSTMENT As String = "Adjustment"

' ------------------------------------------------------------------ reasons
Public Const REASON_USED_UP As String = "Used Up"
Public Const REASON_EXPIRED As String = "Expired"
Public Const REASON_DAMAGED As String = "Damaged"
Public Const REASON_MISSING As String = "Missing"
Public Const REASON_CORRECTION As String = "Correction"
Public Const REASON_OTHER As String = "Other"

' ------------------------------------------------------------------ message classes
Public Enum MsgClass
    mcBlocking = 1
    mcConfirm = 2
    mcInfo = 3
    mcError = 4
End Enum

' ------------------------------------------------------------------ ID formats
Public Const FMT_CONTAINER_ID As String = "C000000"
Public Const FMT_PRODUCT_ID As String = "P000000"
Public Const FMT_TRANSACTION_ID As String = "T00000000"
Public Const FMT_SUPPLIER_ID As String = "S000000"
Public Const FMT_LOCATION_ID As String = "LOC0000"
Public Const FMT_BARCODE As String = "0000000"
Public Const BARCODE_PATTERN As String = "\d{7}"

' ------------------------------------------------------------------ misc
Public Const OPERATOR_UNKNOWN As String = "UNKNOWN"
Public Const PREV_NONE As String = "(none)"
Public Const DEFAULT_QUANTITY As Long = 1
