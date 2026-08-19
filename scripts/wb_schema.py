"""Shared schema constants for the Laboratory Inventory v0.1 workbook.

Single source of truth for sheet/table/column/named-range definitions used by
the builder and by the structural inspection/tests. Keeping these in Python
lets the builder and the tests agree without parsing the YAML contract.

IMPORTANT (Excel-runtime fix 2026-08-17): the header text of every Excel Table
column MUST equal its column name exactly (PascalCase, no spaces). Excel
resolves structured references like tblContainers[ExpiryDate] by the header
cell text; a header "Expiry Date" makes the reference #REF!. The contract
already defines PascalCase column names, so headers now mirror them exactly.
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
# Smallest practical v1 status set (D-016): Reserved removed — no complete
# reservation workflow exists in v1, and a state only enterable via Adjustment
# must not be retained.
STATUSES = ["Available", "InUse", "Expired", "Damaged", "Disposed", "Missing"]
TRANSACTION_TYPES = [
    "Receive", "TakeOpen", "Return", "Transfer", "Dispose",
    "MarkExpired", "MarkDamaged", "MarkMissing", "Adjustment",
]
EXPIRY_CLASSES = ["Expired", "ExpiringSoon", "Valid", "NoExpiry", "Invalid"]
LOCATION_TYPES = ["Cabinet", "Shelf", "Fridge", "Freezer", "Room", "Other"]
DISPOSAL_REASONS = ["Used Up", "Expired", "Damaged", "Missing", "Other"]
TRANSACTION_REASONS = ["Used Up", "Expired", "Damaged", "Missing", "Correction", "Other"]
BOOLEANS = ["TRUE", "FALSE"]


def _cols(*items):
    """Build column tuples (name, header, width) with header == name.

    Header text must equal the column name so Excel structured references
    resolve exactly (tblContainers[ExpiryDate] needs a header 'ExpiryDate').
    """
    return [(name, name, width) for (name, width) in items]


# ---------------------------------------------------------------- tables
# Each entry: name -> dict(sheet, key, start_col, columns=[(name, header, width)])
TABLES = {
    "tblProducts": {
        "sheet": "Products",
        "key": "ProductID",
        "start_col": 1,
        "columns": _cols(
            ("ProductID", 11),
            ("ProductName", 30),
            ("ProductType", 13),
            ("Category", 13),
            ("Manufacturer", 20),
            ("ManufacturerCatalogueNumber", 16),
            ("CASNumber", 14),
            ("Concentration", 12),
            ("Grade", 12),
            ("StandardContainerDescription", 20),
            ("SupplierID", 12),
            ("StorageRequirements", 18),
            ("HazardClassification", 16),
            ("SDSReference", 14),
            ("MinimumContainerStock", 12),
            ("TargetContainerStock", 12),
            ("ReorderQuantity", 12),
            ("Active", 8),
            ("Notes", 24),
            ("HelperAvailableStock", 14),
            ("HelperStockClass", 12),
        ),
        # calculated/derived columns: never manually edited; protected in the
        # workbook; excluded from authoritative data entry (D-017)
        "calculated_columns": [
            "HelperAvailableStock",
            "HelperStockClass",
        ],
    },
    "tblContainers": {
        "sheet": "Containers",
        "key": "ContainerID",
        "start_col": 1,
        "columns": _cols(
            ("ContainerID", 12),
            ("Barcode", 11),
            ("ProductID", 11),
            ("BatchLotNumber", 14),
            ("ExpiryDate", 12),
            ("RetestDate", 12),
            ("DateReceived", 12),
            ("StorageLocationID", 12),
            ("Status", 11),
            ("OpenedDate", 12),
            ("DisposalDate", 12),
            ("DisposalReason", 14),
            ("Notes", 24),
            ("HelperContainerNum", 12),
            ("HelperBarcodeNum", 12),
        ),
        # calculated/derived columns: numeric mirrors of ContainerID/Barcode
        # used by MAX() for next-ID generation; protected (D-017)
        "calculated_columns": [
            "HelperContainerNum",
            "HelperBarcodeNum",
        ],
    },
    "tblTransactions": {
        "sheet": "Transactions",
        "key": "TransactionID",
        "start_col": 1,
        "columns": _cols(
            ("TransactionID", 12),
            ("Timestamp", 17),
            ("Operator", 12),
            ("Barcode", 11),
            ("ContainerID", 12),
            ("ProductID", 11),
            ("ProductName", 28),
            ("TransactionType", 14),
            ("PreviousStatus", 14),
            ("NewStatus", 12),
            ("PreviousLocation", 14),
            ("NewLocation", 14),
            ("BatchLotNumber", 14),
            ("Reason", 14),
            ("Reference", 12),
            ("Notes", 24),
        ),
    },
    "tblSuppliers": {
        "sheet": "Suppliers",
        "key": "SupplierID",
        "start_col": 1,
        "columns": _cols(
            ("SupplierID", 12),
            ("SupplierName", 24),
            ("ContactName", 18),
            ("Email", 22),
            ("Phone", 14),
            ("Website", 20),
            ("Address", 24),
            ("Notes", 20),
        ),
    },
    "tblLocations": {
        "sheet": "Locations",
        "key": "StorageLocationID",
        "start_col": 1,
        "columns": _cols(
            ("StorageLocationID", 14),
            ("LocationName", 22),
            ("LocationType", 14),
            ("Description", 22),
            ("Active", 8),
        ),
    },
    "tblSettings": {
        "sheet": "Settings",
        "key": "SettingKey",
        "start_col": 1,
        "columns": _cols(
            ("SettingKey", 24),
            ("SettingValue", 16),
            ("Description", 40),
        ),
    },
    "tblStatusList": {
        "sheet": "Settings",
        "key": "StatusValue",
        "start_col": 1,
        "columns": _cols(
            ("StatusValue", 16),
            ("Label", 24),
        ),
    },
    "tblTransactionTypeList": {
        "sheet": "Settings",
        "key": "TransactionTypeValue",
        "start_col": 1,
        "columns": _cols(
            ("TransactionTypeValue", 20),
            ("Label", 24),
        ),
    },
    "tblExpiryClassList": {
        "sheet": "Settings",
        "key": "ExpiryClassValue",
        "start_col": 1,
        "columns": _cols(
            ("ExpiryClassValue", 18),
            ("Label", 24),
        ),
    },
    "tblScanResults": {
        "sheet": "Scan",
        "key": "Barcode",
        "start_col": 4,
        "authoritative": False,
        "columns": _cols(
            ("Barcode", 11),
            ("ContainerID", 12),
            ("ProductID", 11),
            ("ProductName", 26),
            ("BatchLotNumber", 14),
            ("ExpiryDate", 12),
            ("StorageLocationID", 12),
            ("LocationName", 18),
            ("Status", 11),
            ("OpenedDate", 12),
            ("ExpiryClass", 13),
            ("DuplicateFlag", 12),
            ("LookupState", 12),
        ),
    },
    "tblReceiveStaging": {
        "sheet": "Receiving",
        "key": "ProductID",
        "start_col": 2,
        "authoritative": False,
        "columns": _cols(
            ("ProductID", 11),
            ("ProductName", 26),
            ("NextContainerID", 12),
            ("NextBarcode", 11),
            ("BatchLotNumber", 14),
            ("ExpiryDate", 12),
            ("RetestDate", 12),
            ("StorageLocationID", 12),
            ("Status", 11),
            ("Quantity", 10),
            ("ValidationMessage", 34),
            ("ReadinessState", 14),
        ),
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
