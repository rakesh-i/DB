from pymongo import MongoClient

client = MongoClient("mongodb://localhost:27017/")


db = client['Stock']
db.create_collection("CKN", validator={
  "$jsonSchema": {
    "bsonType": "object",
    "required": [
      "Month",
      "OSBStock",
      "Purchase",
      "Sales",
      "Production",
      "CSStock"
    ],
    "properties": {
      "Month": {
        "bsonType": "date",
        "description": "must be a date and is required"
      },
      "OSBStock": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a numeric value and is required"
      },
      "Purchase": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a numeric value and is required"
      },
      "Sales": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a numeric value and is required"
      },
      "Production": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a numeric value and is required"
      },
      "CSStock": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a numeric value and is required"
      }
    }
}
})
col = db["CKN"]
col.create_index([('Month',1)], unique=True)

db.create_collection("RCN", validator={
  "$jsonSchema": {
    "bsonType": "object",
    "required": [
      "Month",
      "OSBStock",
      "Purchase",
      "Sales",
      "Production",
      "CSStock"
    ],
    "properties": {
      "Month": {
        "bsonType": "date",
        "description": "must be a date and is required"
      },
      "OSBStock": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a numeric value and is required"
      },
      "Purchase": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a numeric value and is required"
      },
      "Sales": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a numeric value and is required"
      },
      "Production": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a numeric value and is required"
      },
      "CSStock": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a numeric value and is required"
      }
    }
  }

})
col = db["RCN"]
col.create_index([('Month',1)], unique=True)

db.create_collection("Production", validator={
  "$jsonSchema": {
    "bsonType": "object",
    "required": [
      "Lot",
      "Date",
      "RCNroast"
    ],
    "properties": {
      "Lot": {
        "bsonType": "int",
        "description": "Lot number of the roast"
      },
      "Date": {
        "bsonType": "date",
        "description": "Date of the roast"
      },
      "RCNroast": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "RCN roast value"
      }
    }
  }
}
)
col = db["Production"]
col.create_index([('Lot',1)], unique=True)

db.create_collection("Purchase", validator={
  "$jsonSchema": {
    "bsonType": "object",
    "required": [
      "BillNo",
      "Date",
      "Kernels",
      "Shells",
      "RCN",
      "Husk",
      "Dust"
    ],
    "properties": {
      "BillNo": {
        "bsonType": "int",
        "description": "must be an integer and is required"
      },
      "Date": {
        "bsonType": "date",
        "description": "must be a date and is required"
      },
      "Kernels": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a number and is required"
      },
      "Shells": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a number and is required"
      },
      "RCN": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a number and is required"
      },
      "Husk": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a number and is required"
      },
      "Dust": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a number and is required"
      }
    }
  }
}
)
col = db["Purchase"]
col.create_index([('BillNo',1)], unique=True)

db.create_collection("Sales", validator={
  "$jsonSchema": {
    "bsonType": "object",
    "required": [
      "BillNo",
      "Date",
      "Kernels",
      "Shells",
      "RCN",
      "Husk",
      "Dust"
    ],
    "properties": {
      "BillNo": {
        "bsonType": "int",
        "description": "must be an integer and is required"
      },
      "Date": {
        "bsonType": "date",
        "description": "must be a date and is required"
      },
      "Kernels": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a number and is required"
      },
      "Shells": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a number and is required"
      },
      "RCN": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a number and is required"
      },
      "Husk": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a number and is required"
      },
      "Dust": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "must be a number and is required"
      }
    }
  }
}
)
col = db["Sales"]
col.create_index([('BillNo',1)], unique=True)

db = client['Bills']
db.create_collection("Cash", validator={
  "$jsonSchema": {
    "bsonType": "object",
    "required": [
      "Sl",
      "Date",
      "Party",
      "Rate",
      "Qty",
      "QtyType"
    ],
    "properties": {
      "Sl": {
        "bsonType": "int"
      },
      "Date": {
        "bsonType": "date",
        "description": "Date of transaction"
      },
      "Rate": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "Invoice number"
      },
      "Party": {
        "bsonType": "string",
        "description": "Name of the party"
      },
      "Qty": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "Quantity of the item"
      },
      "QtyType": {
        "bsonType": "string",
        "description": "Type of quantity (e.g. kg, lbs)"
      }
    }
  }
}
)
col = db["Cash"]
col.create_index([('Sl',1)], unique=True)

db.create_collection("GSTApp", validator={
  "$jsonSchema": {
    "bsonType": "object",
    "required": [
      "Date",
      "scgst",
      "ssgst",
      "sigst",
      "pcgst",
      "psgst",
      "pigst",
      "Paid",
      "PDate"
    ],
    "properties": {
      "Date": {
        "bsonType": "date",
        "description": "Date field for the collection"
      },
      "scgst": {
        "bsonType": [
          "double",
          "int"
        ],
        "description": "CGST for sales"
      },
      "ssgst": {
        "bsonType": [
          "double",
          "int"
        ],
        "description": "SGST for sales"
      },
      "sigst": {
        "bsonType": [
          "double",
          "int"
        ],
        "description": "IGST for sales"
      },
      "pcgst": {
        "bsonType": [
          "double",
          "int"
        ],
        "description": "CGST for purchase"
      },
      "psgst": {
        "bsonType": [
          "double",
          "int"
        ],
        "description": "SGST for purchase"
      },
      "pigst": {
        "bsonType": [
          "double",
          "int"
        ],
        "description": "IGST for purchase"
      },
      "Paid": {
        "bsonType": [
          "double",
          "int"
        ],
        "description": "Amount paid"
      },
      "PDate": {
        "bsonType": "date",
        "description": "Date when paid"
      }
    }
  }
}
)
col = db["GSTApp"]
col.create_index([('Date',1)], unique=True)

db.create_collection('PSdetails', validator={
  "$jsonSchema": {
    "bsonType": "object",
    "required": [
      "Date",
      "Party",
      "INVNo",
      "Qty",
      "QtyType",
      "Amount",
      "SGST",
      "CGST",
      "IGST"
    ],
    "properties": {
      "Date": {
        "bsonType": "date",
        "description": "Date of transaction"
      },
      "INVNo": {
        "bsonType": "string",
        "description": "Invoice number"
      },
      "Party": {
        "bsonType": "string",
        "description": "Name of the party"
      },
      "Qty": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "Quantity of the item"
      },
      "QtyType": {
        "bsonType": "string",
        "description": "Type of quantity (e.g. kg, lbs)"
      },
      "Amount": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "Amount of the transaction"
      },
      "SGST": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "Amount of SGST"
      },
      "CGST": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "Amount of CGST"
      },
      "IGST": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "Amount of IGST"
      }
    }
  }
}
)
col = db["PSdetails"]
col.create_index([('Date',1)], unique=True)

db.create_collection('Sdetails', validator={
  "$jsonSchema": {
    "bsonType": "object",
    "required": [
      "Date",
      "Party",
      "INVNo",
      "Qty",
      "QtyType",
      "Amount",
      "SGST",
      "CGST",
      "IGST"
    ],
    "properties": {
      "Date": {
        "bsonType": "date",
        "description": "Date of transaction"
      },
      "INVNo": {
        "bsonType": "string",
        "description": "Invoice number"
      },
      "Party": {
        "bsonType": "string",
        "description": "Name of the party"
      },
      "Qty": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "Quantity of the item"
      },
      "QtyType": {
        "bsonType": "string",
        "description": "Type of quantity (e.g. kg, lbs)"
      },
      "Amount": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "Amount of the transaction"
      },
      "SGST": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "Amount of SGST"
      },
      "CGST": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "Amount of CGST"
      },
      "IGST": {
        "bsonType": [
          "int",
          "double"
        ],
        "description": "Amount of IGST"
      }
    }
  }
}
)
col = db["Sdetails"]
col.create_index([('Date',1)], unique=True)

db = client['GST']
db.create_collection("Purchase", validator={
  "$jsonSchema": {
    "bsonType": "object",
    "required": [
      "Id",
      "Date",
      "Party",
      "VoucherNo",
      "GSTIN_UIN",
      "Qty",
      "Unit",
      "Purchase",
      "CGST",
      "SGST",
      "IGST",
      "RoundOff",
      "Total"
    ],
    "properties": {
      "Id": {
        "bsonType": "int"
      },
      "Date": {
        "bsonType": "date"
      },
      "Party": {
        "bsonType": "string"
      },
      "VoucherNo": {
        "bsonType": "string"
      },
      "GSTIN_UIN": {
        "bsonType": "string"
      },
      "Qty": {
        "bsonType": [
          "double",
          "int"
        ]
      },
      "Unit": {
        "bsonType": "string"
      },
      "Purchase": {
        "bsonType": [
          "double",
          "int"
        ]
      },
      "CGST": {
        "bsonType": [
          "double",
          "int"
        ]
      },
      "SGST": {
        "bsonType": [
          "double",
          "int"
        ]
      },
      "IGST": {
        "bsonType": [
          "double",
          "int"
        ]
      },
      "Roundoff": {
        "bsonType": [
          "double",
          "int"
        ]
      },
      "Total": {
        "bsonType": [
          "double",
          "int"
        ]
      }
    }
  }
}
)
col = db["Purchase"]
col.create_index([('Id',1)], unique=True)

db.create_collection("Sales", validator={
  "$jsonSchema": {
    "bsonType": "object",
    "required": [
      "Id",
      "Date",
      "Party",
      "VoucherNo",
      "GSTIN_UIN",
      "Qty",
      "Unit",
      "Sales",
      "CGST",
      "SGST",
      "IGST",
      "RoundOff",
      "Total"
    ],
    "properties": {
      "Id": {
        "bsonType": "int"
      },
      "Date": {
        "bsonType": "date"
      },
      "Party": {
        "bsonType": "string"
      },
      "VoucherNo": {
        "bsonType": "string"
      },
      "GSTIN_UIN": {
        "bsonType": "string"
      },
      "Qty": {
        "bsonType": [
          "double",
          "int"
        ]
      },
      "Unit": {
        "bsonType": "string"
      },
      "Sales": {
        "bsonType": [
          "double",
          "int"
        ]
      },
      "CGST": {
        "bsonType": [
          "double",
          "int"
        ]
      },
      "SGST": {
        "bsonType": [
          "double",
          "int"
        ]
      },
      "IGST": {
        "bsonType": [
          "double",
          "int"
        ]
      },
      "Roundoff": {
        "bsonType": [
          "double",
          "int"
        ]
      },
      "Total": {
        "bsonType": [
          "double",
          "int"
        ]
      }
    }
  }
}
)
col = db["Sales"]
col.create_index([('Id',1)], unique=True)

db = client["Container"]
db.create_collection("Total", validator={
      "$jsonSchema": {
        "bsonType": "object",
        "required": ["Con", "Details"],
        "properties": {
          "Con": {
            "bsonType": "int",
            "description": "must be an integer and is required"
          },
          "Details": {
            "bsonType": "string",
            "description": "must be a string and is required"
          }
        }
      }
    }
    )
col = db["Total"]
col.create_index(['Con', 1], unique=True)

db.create_collection("Wholes", validator={
      "$jsonSchema": {
        "bsonType": "object",
        "required": [
          "Con",
          "Grade",
          "GradeD",
          "Trip1",
          "Trip2",
          "Trip3",
          "Trip4",
          "Trip5",
          "Trip6",
          "Trip7",
          "Total"
        ],
        "properties": {
          "Con": {
            "bsonType": "int"
          },
          "Grade": {
            "bsonType": "string"
          },
          "GradeD": {
            "bsonType": "string"
          },
          "Trip1": {
            "bsonType": [
              "double",
              "int"
            ]
          },
          "Trip2": {
            "bsonType": [
              "double",
              "int"
            ]
          },
          "Trip3": {
            "bsonType": [
              "double",
              "int"
            ]
          },
          "Trip4": {
            "bsonType": [
              "double",
              "int"
            ]
          },
          "Trip5": {
            "bsonType": [
              "double",
              "int"
            ]
          },
          "Trip6": {
            "bsonType": [
              "double",
              "int"
            ]
          },
          "Trip7": {
            "bsonType": [
              "double",
              "int"
            ]
          },
          "Total": {
            "bsonType": [
              "double",
              "int"
            ]
          }
        }
      }
    }
    )
col = db["Wholes"]
col.create_index(['Con', 1], unique=True)

db.create_collection("Pieces", validator={
      "$jsonSchema": {
        "bsonType": "object",
        "required": [
          "Con",
          "Grade",
          "GradeD",
          "Trip1",
          "Trip2",
          "Trip3",
          "Trip4",
          "Trip5",
          "Trip6",
          "Trip7",
          "Total"
        ],
        "properties": {
          "Con": {
            "bsonType": "int"
          },
          "Grade": {
            "bsonType": "string"
          },
          "GradeD": {
            "bsonType": "string"
          },
          "Trip1": {
            "bsonType": [
              "double",
              "int"
            ]
          },
          "Trip2": {
            "bsonType": [
              "double",
              "int"
            ]
          },
          "Trip3": {
            "bsonType": [
              "double",
              "int"
            ]
          },
          "Trip4": {
            "bsonType": [
              "double",
              "int"
            ]
          },
          "Trip5": {
            "bsonType": [
              "double",
              "int"
            ]
          },
          "Trip6": {
            "bsonType": [
              "double",
              "int"
            ]
          },
          "Trip7": {
            "bsonType": [
              "double",
              "int"
            ]
          },
          "Total": {
            "bsonType": [
              "double",
              "int"
            ]
          }
        }
      }
    }
    )
col = db["Pieces"]
col.create_index(['Con', 1], unique=True)

