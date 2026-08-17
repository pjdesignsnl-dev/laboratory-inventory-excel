"""Shared schema constants for the Laboratory Inventory v0.1 workbook.

Single source of truth for sheet/table/column/named-range definitions used by
the builder and by the structural inspection/tests. Keeping these in Python
lets the builder and the tests agree without parsing the YAML contract.
"""
from __future__ import annotations

# ---------------------------------------------------------------- sheets
SHEET_ORDER = [
    "Dashboard",
    "Scan",
    "Receiving",
    "Products",
    "Containers",
    "Transactions",
    "Suppliers",
    "Locations",
    "Settings",
]

TAB_COLORS = {
    "Dashboard": "1F4E79",
    "Scan": "2E7D32",
    "Receiving": "00838F",
    "Products": "1565C0",
    "Containers": "283593",
    "Transactions": "6A1B9A",
    "Suppliers": "455A64",
    "Locations": "37474F",
    "Settings": "78909C",
}

# ---------------------------------------------------------------- lists
PRODUCT_TYPES = ["Consumable", "Chemical", "Reagent"]
CATEGORIES = ["General", "Pipette Tips", "Tubes", "Solvent", "Reagent", "Consumable"]
STATUSES = ["Available", "InUse", "Reserved", "Expired", "Damaged", "Disposed", "Missing"]
TRANSACTION_TYPES = [
    "Receive", "TakeOpen", "Return", "Transfer", "Dispose",
    "MarkExpired", "MarkDamaged", "MarkMissing", "Adjustment",
]
EXPIRY_CLASSES = ["Expired", "ExpiringSoon", "Valid", "NoExpiry", "Invalid"]
LOCATION_TYPES = ["Cabinet", "Shelf", "Fridge", "Freezer", "Room", "Other"]
DISPOSAL_REASONS = ["Used Up", "Expired", "Damaged", "Missing", "Other"]
TRANSACTION_REASONS = ["Used Up", "Expired", "Damaged", "Missing", "Correction", "Other"]
BOOLEANS = ["TRUE", "FALSE"]

# ---------------------------------------------------------------- tables
# Each entry: name -> dict(sheet, key, start_col, columns=[(name, header, width)])
TABLES = {
    "tblProducts": {
        "sheet": "Products",
        "key": "ProductID",
        "start_col": 1,
        "columns": [
            ("ProductID", "Product ID", 11),
            ("ProductName", "Product Name", 30),
            ("ProductType", "Product Type", 13),
            ("Category", "Category", 13),
            ("Manufacturer", "Manufacturer", 20),
            ("ManufacturerCatalogueNumber", "Manufacturer Catalogue Number", 16),
            ("CASNumber", "CAS Number", 14),
            ("Concentration", "Concentration", 12),
            ("Grade", "Grade / Purity", 12),
            ("StandardContainerDescription", "Standard Container Description", 20),
            ("SupplierID", "Supplier ID", 12),
            ("StorageRequirements", "Storage Requirements", 18),
            ("HazardClassification", "Hazard Classification", 16),
            ("SDSReference", "SDS Reference", 14),
            ("MinimumContainerStock", "Minimum Container Stock", 12),
            ("TargetContainerStock", "Target Container Stock", 12),
            ("ReorderQuantity", "Reorder Quantity", 12),
            ("Active", "Active", 8),
            ("Notes", "Notes", 24),
            ("HelperAvailableStock", "Available Stock (formula)", 14),
            ("HelperStockClass", "Stock Class (formula)", 12),
        ],
    },
    "tblContainers": {
        "sheet": "Containers",
        "key": "ContainerID",
        "start_col": 1,
        "columns": [
            ("ContainerID", "Container ID", 12),
            ("Barcode", "Barcode", 11),
            ("ProductID", "Product ID", 11),
            ("BatchLotNumber", "Batch / Lot Number", 14),
            ("ExpiryDate", "Expiry Date", 12),
            ("RetestDate", "Retest Date", 12),
            ("DateReceived", "Date Received", 12),
            ("StorageLocationID", "Storage Location ID", 12),
            ("Status", "Status", 11),
            ("OpenedDate", "Opened Date", 12),
            ("DisposalDate", "Disposal Date", 12),
            ("DisposalReason", "Disposal Reason", 14),
            ("Notes", "Notes", 24),
            ("HelperContainerNum", "Helper Container Num", 12),
            ("HelperBarcodeNum", "Helper Barcode Num", 12),
        ],
    },
    "tblTransactions": {
        "sheet": "Transactions",
        "key": "TransactionID",
        "start_col": 1,
        "columns": [
            ("TransactionID", "Transaction ID", 12),
            ("Timestamp", "Timestamp", 17),
            ("Operator", "Operator", 12),
            ("Barcode", "Barcode", 11),
            ("ContainerID", "Container ID", 12),
            ("ProductID", "Product ID", 11),
            ("ProductName", "Product Name", 28),
            ("TransactionType", "Transaction Type", 14),
            ("PreviousStatus", "Previous Status", 14),
            ("NewStatus", "New Status", 12),
            ("PreviousLocation", "Previous Location", 14),
            ("NewLocation", "New Location", 14),
            ("BatchLotNumber", "Batch / Lot Number", 14),
            ("Reason", "Reason", 14),
            ("Reference", "Reference", 12),
            ("Notes", "Notes", 24),
        ],
    },
    "tblSuppliers": {
        "sheet": "Suppliers",
        "key": "SupplierID",
        "start_col": 1,
        "columns": [
            ("SupplierID", "Supplier ID", 12),
            ("SupplierName", "Supplier Name", 24),
            ("ContactName", "Contact Name", 18),
            ("Email", "Email", 22),
            ("Phone", "Phone", 14),
            ("Website", "Website", 20),
            ("Address", "Address", 24),
            ("Notes", "Notes", 20),
        ],
    },
    "tblLocations": {
        "sheet": "Locations",
        "key": "StorageLocationID",
        "start_col": 1,
        "columns": [
            ("StorageLocationID", "Storage Location ID", 14),
            ("LocationName", "Location Name", 22),
            ("LocationType", "Location Type", 14),
            ("Description", "Description", 22),
            ("Active", "Active", 8),
        ],
    },
    "tblSettings": {
        "sheet": "Settings",
        "key": "SettingKey",
        "start_col": 1,
        "columns": [
            ("SettingKey", "Setting Key", 24),
            ("SettingValue", "Setting Value", 16),
            ("Description", "Description", 40),
        ],
    },
    "tblStatusList": {
        "sheet": "Settings",
        "key": "StatusValue",
        "start_col": 1,
        "columns": [
            ("StatusValue", "Status Value", 16),
            ("Label", "Label", 24),
        ],
    },
    "tblTransactionTypeList": {
        "sheet": "Settings",
        "key": "TransactionTypeValue",
        "start_col": 1,
        "columns": [
            ("TransactionTypeValue", "Transaction Type Value", 20),
            ("Label", "Label", 24),
        ],
    },
    "tblExpiryClassList": {
        "sheet": "Settings",
        "key": "ExpiryClassValue",
        "start_col": 1,
        "columns": [
            ("ExpiryClassValue", "Expiry Class Value", 18),
            ("Label", "Label", 24),
        ],
    },
    "tblScanResults": {
        "sheet": "Scan",
        "key": "Barcode",
        "start_col": 4,
        "authoritative": False,
        "columns": [
            ("Barcode", "Barcode", 11),
            ("ContainerID", "Container ID", 12),
            ("ProductID", "Product ID", 11),
            ("ProductName", "Product Name", 26),
            ("BatchLotNumber", "Batch / Lot Number", 14),
            ("ExpiryDate", "Expiry Date", 12),
            ("StorageLocationID", "Storage Location ID", 12),
            ("LocationName", "Location Name", 18),
            ("Status", "Status", 11),
            ("OpenedDate", "Opened Date", 12),
            ("ExpiryClass", "Expiry Class", 13),
            ("DuplicateFlag", "Duplicate Flag", 12),
            ("LookupState", "Lookup State", 12),
        ],
    },
    "tblReceiveStaging": {
        "sheet": "Receiving",
        "key": "ProductID",
        "start_col": 2,
        "authoritative": False,
        "columns": [
            ("ProductID", "Product ID", 11),
            ("ProductName", "Product Name", 26),
            ("NextContainerID", "Next Container ID", 12),
            ("NextBarcode", "Next Barcode", 11),
            ("BatchLotNumber", "Batch / Lot Number", 14),
            ("ExpiryDate", "Expiry Date", 12),
            ("RetestDate", "Retest Date", 12),
            ("StorageLocationID", "Storage Location ID", 12),
            ("Status", "Status", 11),
            ("Quantity", "Quantity", 10),
            ("ValidationMessage", "Validation Message", 34),
            ("ReadinessState", "Readiness State", 14),
        ],
    },
}

# ---------------------------------------------------------------- named ranges
# Interface cell named ranges (fixed for contract). Layout convention:
# labels in column B, inputs/outputs in column D on Scan/Receiving.
NAMED_RANGES = {
    "rngScanInput": ("Scan", "D7"),
    "rngScanStatusMessage": ("Scan", "D9"),
    "rngScanResultCard": ("Scan", "D4:K34"),
    "rngReceiveProductID": ("Receiving", "D7"),
    "rngReceiveNextContainerID": ("Receiving", "D9"),
    "rngReceiveNextBarcode": ("Receiving", "D10"),
    "rngReceiveLot": ("Receiving", "D12"),
    "rngReceiveExpiry": ("Receiving", "D13"),
    "rngReceiveRetest": ("Receiving", "D14"),
    "rngReceiveLocation": ("Receiving", "D15"),
    "rngReceiveQuantity": ("Receiving", "D16"),
    "rngReceiveStatusMessage": ("Receiving", "D19"),
}

# Settings-key workbook names -> Settings sheet value cell
SETTINGS_NAMES = {
    "ExpiryWarningDays30": ("Settings", "B6"),
    "ExpiryWarningDays60": ("Settings", "B7"),
    "ExpiryWarningDays90": ("Settings", "B8"),
    "DefaultLocationID": ("Settings", "B9"),
    "DefaultStatusNewContainers": ("Settings", "B10"),
}

# Column-list named ranges (structured reference lists for validation/lookup)
COLUMN_NAMES = {
    "lstProductsProductID": ("Products", "A", "tblProducts[ProductID]"),
    "lstProductsProductName": ("Products", "B", "tblProducts[ProductName]"),
    "lstStatusList": ("Settings", "A", "tblStatusList[StatusValue]"),
    "lstTransactionTypeList": ("Settings", "A", "tblTransactionTypeList[TransactionTypeValue]"),
    "lstExpiryClassList": ("Settings", "A", "tblExpiryClassList[ExpiryClassValue]"),
    "lstSupplierIDs": ("Suppliers", "A", "tblSuppliers[SupplierID]"),
    "lstLocationIDs": ("Locations", "A", "tblLocations[StorageLocationID]"),
}

LIST_RANGES = {
    "lstProductType": ("Settings", "H4:H6", PRODUCT_TYPES),
    "lstCategory": ("Settings", "H8:H13", CATEGORIES),
    "lstLocationType": ("Settings", "H15:H20", LOCATION_TYPES),
    "lstDisposalReason": ("Settings", "H22:H26", DISPOSAL_REASONS),
    "lstTransactionReason": ("Settings", "H28:H33", TRANSACTION_REASONS),
    "lstBool": ("Settings", "H35:H36", BOOLEANS),
}

# Settings keys (SettingKey -> (value, description)) in display order.
# B-column rows: key at row 6..12 -> value cells B6..B12.
SETTINGS = {
    "ExpiryWarningDays30": ("30", "Expiry warning band: days until expiry to flag (30)"),
    "ExpiryWarningDays60": ("60", "Expiry warning band: days until expiry to flag (60)"),
    "ExpiryWarningDays90": ("90", "Expiry warning band: days until expiry to flag (90)"),
    "DefaultLocationID": ("LOC0001", "Default storage location for new containers"),
    "DefaultStatusNewContainers": ("Available", "Status assigned to newly received containers"),
    "ScannerEnterSuffix": ("TRUE", "Barcode scanner configured to send Enter suffix"),
    "WorkbookVersion": ("v0.1", "Workbook version label"),
}

WORKBOOK_VERSION = "v0.1"
CONTRACT_VERSION = "0.1.0"
