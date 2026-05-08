const firstDataRow = 7;

function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    if (!usedRange) return;
    const values = usedRange.getValues();
    const ctx: Context = {
        headers: values[3],
        productType: sheet.getName().replace("Vorlage-", ""),
        barcodes: values.slice(firstDataRow - 1).map(row => row[10])
    };

    for (const rule of Rules) {
        if (
            rule.appliesTo.includes(ctx.productType) ||
            rule.appliesTo.includes("ALL")
        ){
            fill(sheet, values, ctx, rule.headerName, rule.value)
        }
    }

    usedRange.setValues(values);
}

//////////////////////////////////////////
//Types
type CellValue = string | number | boolean | null;
type Context = {
    headers: CellValue[];
    productType: string;
    barcodes: CellValue[];
}
type Rule = {
    headerName: string; 
    value: string | number | ((ctx: Context, row: number) => string | number);
    appliesTo: string[];
};

//////////////////////////////////////////
//Rules

const Rules: Rule[] = [
    {
        headerName: "rtip_vendor_code#1.value",
        value: "heo GmbH, de_toys, HEOGM",
        appliesTo: ["ALL"]
    },
    {
        headerName: "product_type#1.value",
        value: (ctx: Context, row: number) => ctx.productType,
        appliesTo: ["ALL"]
    },
    {
        headerName: "external_product_id#1.type",
        value: (ctx: Context, row: number) => {
            const raw = ctx.barcodes[row - (firstDataRow - 1)];
            if (raw == null) {return "ERROR";}
            const barcode = String(raw).trim();
            if (barcode.length === 12) return "UPC";
            if (barcode.length === 13) return "EAN";
            return "ERROR";
        },
        appliesTo: ["ALL"]
    },
    {
        headerName: "package_level#1.value",
        value: "Einheit",
        appliesTo: ["ALL"]
    },
    {
        headerName: "is_trade_item_orderable_unit#1.value",
        value: "Ja",
        appliesTo: ["ALL"]
    },
    {
        headerName: "cost_price#1.currency",
        value: "EUR",
        appliesTo: ["ALL"]
    },
    {
        headerName: "item_package_dimensions#1.length.unit",
        value: "Millimeter",
        appliesTo: ["ALL"]
    },
    {
        headerName: "item_package_dimensions#1.width.unit",
        value: "Millimeter",
        appliesTo: ["ALL"]
    },
    {
        headerName: "item_package_dimensions#1.height.unit",
        value: "Millimeter",
        appliesTo: ["ALL"]
    },
    {
        headerName: "item_package_weight#1.unit",
        value: "Gramm",
        appliesTo: ["ALL"]
    },
    {
        headerName: "rtip_items_per_inner_pack#1.value",
        value: 1,
        appliesTo: ["ALL"]
    },
    {
        headerName: "is_assembly_required#1.value",
        value: "Nein",
        appliesTo: ["TOY_FIGURE"]
    },
    {
        headerName: "unit_count#1.value",
        value: 1,
        appliesTo: ["TOY_FIGURE", "FIGURINE"]
    },
    {
        headerName: "unit_count#1.type.value",
        value: "stück",
        appliesTo: ["TOY_FIGURE", "FIGURINE"]
    },
    {
        headerName: "street_date#1.value",
        value: (ctx: Context, row: number) => {
            const today = new Date();
            return "'" + today.toISOString().split("T")[0];
        },
        appliesTo: ["TOY_FIGURE"]
    },
    {
        headerName: "manufacturer_minimum_age#1.value",
        value: 168,
        appliesTo: ["TOY_FIGURE"]
    },
    {
        headerName: "manufacturer_maximum_age#1.value",
        value: 1200,
        appliesTo: ["TOY_FIGURE"]
    },
    {
        headerName: "item_length_width_height#1.height.unit",
        value: "Zentimeter",
        appliesTo: ["TOY_FIGURE", "MONEY_BANK", "FIGURINE"]
    },
    {
        headerName: "item_length_width_height#1.length.unit",
        value: "Zentimeter",
        appliesTo: ["TOY_FIGURE", "MONEY_BANK", "FIGURINE"]
    },
    {
        headerName: "item_length_width_height#1.width.unit",
        value: "Zentimeter",
        appliesTo: ["TOY_FIGURE", "MONEY_BANK", "FIGURINE"]
    },
    {
        headerName: "warranty_description#1.value",
        value: "Keine",
        appliesTo: ["TOY_FIGURE"]
    },
    {
        headerName: "rtip_safety_warning#1.value",
        value: "Keine",
        appliesTo: ["TOY_FIGURE"]
    },
    {
        headerName: "batteries_required#1.value",
        value: "Nein",
        appliesTo: ["TOY_FIGURE", "MONEY_BANK", "FIGURINE"]
    },
    {
        headerName: "eu_toys_safety_directive_language#1.value",
        value: "Englisch",
        appliesTo: ["TOY_FIGURE"]
    },
    {
        headerName: "eu_toys_safety_directive_language#2.value",
        value: "Deutsch",
        appliesTo: ["TOY_FIGURE"]
    },
    {
        headerName: "eu_toys_safety_directive_age_warning#1.value",
        value: "Nicht geeignet für Kinder unter 14 Jahren",
        appliesTo: ["TOY_FIGURE"]
    },
    {
        headerName: "eu_toys_safety_directive_warning#1.value",
        value: "Kein Warnhinweis zutreffend",
        appliesTo: ["TOY_FIGURE"]
    },
    {
        headerName: "number_of_items#1.value",
        value: 1,
        appliesTo: ["MONEY_BANK", "FIGURINE"]
    },
    {
        headerName: "color#1.value",
        value: "Mehrfarbig",
        appliesTo: ["MONEY_BANK"]
    },
    {
        headerName: "color#1.value",
        value: "Mehrfarbig",
        appliesTo: ["MONEY_BANK", "FIGURINE"]
    },
    {
        headerName: "is_fragile#1.value",
        value: "Nein",
        appliesTo: ["MONEY_BANK", "FIGURINE"]
    },
    {
        headerName: "supplier_declared_dg_hz_regulation#1.value",
        value: "Nicht zutreffend",
        appliesTo: ["MONEY_BANK", "FIGURINE"]
    },
];

//////////////////////////////////////////
//Funktionen

function fill(
    sheet: ExcelScript.Worksheet,
    values: (CellValue)[][],
    ctx: (Context),
    headerName: (string),
    value: (string | number | ((ctx: Context, row: number) => string | number))

) {
    let col = -1;
    col = ctx.headers.indexOf(headerName)

    if (col === -1) {
        throw new Error(headerName + " nicht gefunden");
        //return;
    }
    for (let row = firstDataRow - 1; row < values.length; row++) {
        values[row][col] = 
            typeof value === "function"
                ? value(ctx, row)
                : value;
    }
}
