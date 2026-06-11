/**
 * Script-Name: Amazon Bilderlinks HR
 * Version: 0.2
 * Autor: Valentin Gerdon
 * Letzte Änderung: 2026-06-10
 * Zweck:
 * - Das Skript schreibt nicht mehr den gesamten Used Range in `main()` zurück. Stattdessen verändert `fill() nur noch die Zellen, für die tatsächlich neue Werte generiert wurden.
 * 
 * Changelog:
 * 0.1 - Erster Entwurf, Architektur übernommen vom Autofiller
 */

const firstDataRow = 7;

function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    if (!usedRange) return;
    const values = usedRange.getValues();

    const ctx: Context = {
        headers: values[3],
        productType: sheet.getName().replace("Vorlage-", ""),
        sku: values.slice(firstDataRow - 1).map(row => row[values[3].indexOf("vendor_sku#1.value")])
    };

    for (const rule of Rules) {
        if (
            rule.appliesTo.includes(ctx.productType) ||
            rule.appliesTo.includes("ALL")
        ) {
            fill(sheet, values, ctx, rule.headerName, rule.value)
        }
    }
}

//////////////////////////////////////////
//Types
type CellValue = string | number | boolean | null;
type Context = {
    headers: CellValue[];
    productType: string;
    sku: CellValue[];
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
        headerName: "main_product_image_locator#1.media_location",
        value: (ctx: Context, row: number) => {
            const temp = String(ctx.sku[row - (firstDataRow - 1)]).toLowerCase();
            if (temp == null) { return "ERROR" }
            let loc = String("https://www.heomedia.com/img/hr/hr_" + temp + ".jpg");
            return loc;
        },
        appliesTo: ["ALL"]
    },
    {
        headerName: "other_product_image_locator_1#1.media_location",
        value: (ctx: Context, row: number) => {
            const temp = String(ctx.sku[row - (firstDataRow - 1)]).toLowerCase();
            if (temp == null) { return "ERROR" }
            let loc = String("https://www.heomedia.com/img/hr/hr_" + temp + "_a.jpg");
            return loc;
        },
        appliesTo: ["ALL"]
    },
    {
        headerName: "other_product_image_locator_2#1.media_location",
        value: (ctx: Context, row: number) => {
            const temp = String(ctx.sku[row - (firstDataRow - 1)]).toLowerCase();
            if (temp == null) { return "ERROR" }
            let loc = String("https://www.heomedia.com/img/hr/hr_" + temp + "_b.jpg");
            return loc;
        },
        appliesTo: ["ALL"]
    },
    {
        headerName: "other_product_image_locator_3#1.media_location",
        value: (ctx: Context, row: number) => {
            const temp = String(ctx.sku[row - (firstDataRow - 1)]).toLowerCase();
            if (temp == null) { return "ERROR" }
            let loc = String("https://www.heomedia.com/img/hr/hr_" + temp + "_c.jpg");
            return loc;
        },
        appliesTo: ["ALL"]
    },
    {
        headerName: "other_product_image_locator_4#1.media_location",
        value: (ctx: Context, row: number) => {
            const temp = String(ctx.sku[row - (firstDataRow - 1)]).toLowerCase();
            if (temp == null) { return "ERROR" }
            let loc = String("https://www.heomedia.com/img/hr/hr_" + temp + "_d.jpg");
            return loc;
        },
        appliesTo: ["ALL"]
    },
    {
        headerName: "other_product_image_locator_5#1.media_location",
        value: (ctx: Context, row: number) => {
            const temp = String(ctx.sku[row - (firstDataRow - 1)]).toLowerCase();
            if (temp == null) { return "ERROR" }
            let loc = String("https://www.heomedia.com/img/hr/hr_" + temp + "_e.jpg");
            return loc;
        },
        appliesTo: ["ALL"]
    },
    {
        headerName: "other_product_image_locator_6#1.media_location",
        value: (ctx: Context, row: number) => {
            const temp = String(ctx.sku[row - (firstDataRow - 1)]).toLowerCase();
            if (temp == null) { return "ERROR" }
            let loc = String("https://www.heomedia.com/img/hr/hr_" + temp + "_f.jpg");
            return loc;
        },
        appliesTo: ["ALL"]
    },
    {
        headerName: "other_product_image_locator_7#1.media_location",
        value: (ctx: Context, row: number) => {
            const temp = String(ctx.sku[row - (firstDataRow - 1)]).toLowerCase();
            if (temp == null) { return "ERROR" }
            let loc = String("https://www.heomedia.com/img/hr/hr_" + temp + "_g.jpg");
            return loc;
        },
        appliesTo: ["ALL"]
    },
    {
        headerName: "other_product_image_locator_8#1.media_location",
        value: (ctx: Context, row: number) => {
            const temp = String(ctx.sku[row - (firstDataRow - 1)]).toLowerCase();
            if (temp == null) { return "ERROR" }
            let loc = String("https://www.heomedia.com/img/hr/hr_" + temp + "_h.jpg");
            return loc;
        },
        appliesTo: ["ALL"]
    },
];

//////////////////////////////////////////
//Funktionen

function fill(
    sheet: ExcelScript.Worksheet,
    values: CellValue[][],
    ctx: Context,
    headerName: string,
    value: string | number | ((ctx: Context, row: number) => string | number)
) {
    let col = -1;
    col = ctx.headers.indexOf(headerName)

    if (col === -1) {
        throw new Error(headerName + " nicht gefunden");
        //return;
    }

    const startRow = firstDataRow - 1
    const rowCount = values.length - startRow;
    const output: (string | number)[][] = [];

    for (let row = startRow; row < values.length; row++) {
        const newValue: CellValue =
            typeof value === "function"
                ? value(ctx, row)
                : value;

        values[row][col] = newValue;
        output.push([newValue]);
    }

    sheet.getRangeByIndexes(startRow, col, rowCount, 1).setValues(output);
}
