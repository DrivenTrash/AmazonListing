/**
 * Script-Name: Amazon Autofiller
 * Version: 1.6
 * Unterstützte Vorlagen: APPAREL_PIN, ART_CRAFT_KIT, BACKPACK, BED_LINEN, BLANKET, CARRIER_BAG_CASE, DECORATIVE_MAGNET, DESK_PAD, DISHWARE_BOWL, DISPLAY_ALBUM, DOLL_CLOTHING, DRESS, DRINKING_CUP, FIGURINE, HOBBIES, HOME, KEYCHAIN, LAMP, MINIATURE_TOY_BUILDING, MINIATURE_TOY_FURNISHIN, MODEL, MONEY_BANK, MOUSE_PAD, NON_RIDING_TOY_VEHICLE, PILLOW, PRETEND_PLAY_TOY, PUZZLES, STICKER_DECAL, STORAGE_BOX, TABLETOP_GAME, THERMOS, TOTE_BAG, TOWEL, TOY_BUILDING_BLOCK, TOY_FIGURE, TOY_GUN, TOYS_AND_GAMES, WALL_ART
 * Autor: Valentin Gerdon
 * Letzte Änderung: 2026-06-08
 * Zweck:
 * - Erweitert für DRINKING_CUP, DISPLAY_ALBUM, DISCHWARE_BOWL, TOY_BUILDING_BLOCK, WALL_ART, MINIATURE_TOY_BUILDING, TOY_GUN
 * 
 * Changelog:
 * - 1.5 Erweitert Für ART_CRAFT_KIT, HOBBIES, ADVENT_CALENDAR, DRESS, PRETEND_PLAY_TOY, MINIATURE_TOY_FURNISHIN
 * - 1.4 Erweitert für DOLL_CLOTHING, BACKPACK, DESK_PAD, HOME, DECORATIVE_MAGNET, BED_LINEN, TOWEL, PUZZLES, PILLOW, BLANKET, STICKER_DECAL, TOYS_AND_GAMES, LAMP, THERMOS
 * - 1.3 Erweitert für KEYCHAIN, MOUSE_PAD, TABLETOP_GAME, NON_RIDING_TOY_VEHICLE, CARRIER_BAG_CASE, TOTE_BAG
 * - 1.2.1 Alles auf Zentimeter umgestellt, da neuer Export
 * - 1.2: Erweitert für APPAREL_PIN, MODEL. STORAGE_BOX
 * - 1.1: Erweitert für FIGURINE und MONEY_BANK
 * - 1.0: Skalierbare Architektur
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
        barcodes: values.slice(firstDataRow - 1).map(row => row[values[3].indexOf("external_product_id#1.value")])
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
        appliesTo: ["TOY_FIGURE", "MONEY_BANK", "FIGURINE", "MODEL", "STORAGE_BOX", "MOUSE_PAD", "TABLETOP_GAME", "NON_RIDING_TOY_VEHICLE", "DOLL_CLOTHING", "DESK_PAD", "HOME", "DECORATIVE_MAGNET", "BED_LINEN", "TOWEL", "PUZZLES", "PILLOW", "BLANKET", "STICKER_DECAL", "TOYS_AND_GAMES", "LAMP", "THERMOS", "ART_CRAFT_KIT", "HOBBIES", "ADVENT_CALENDAR", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "DRINKING_CUP", "DISPLAY_ALBUM", "DISHWARE_BOWL", "TOY_BUILDING_BLOCK", "WALL_ART", "MINIATURE_TOY_BUILDING", "TOY_GUN"]
    },
    {
        headerName: "is_trade_item_orderable_unit#1.value",
        value: "Ja",
        appliesTo: ["TOY_FIGURE", "MONEY_BANK", "FIGURINE", "MODEL", "STORAGE_BOX", "MOUSE_PAD", "TABLETOP_GAME", "NON_RIDING_TOY_VEHICLE", "DOLL_CLOTHING", "DESK_PAD", "HOME", "DECORATIVE_MAGNET", "BED_LINEN", "TOWEL", "PUZZLES", "PILLOW", "BLANKET", "STICKER_DECAL", "TOYS_AND_GAMES", "LAMP", "THERMOS", "ART_CRAFT_KIT", "HOBBIES", "ADVENT_CALENDAR", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "DRINKING_CUP", "DISPLAY_ALBUM", "DISHWARE_BOWL", "TOY_BUILDING_BLOCK", "WALL_ART", "MINIATURE_TOY_BUILDING", "TOY_GUN"]
    },
    {
        headerName: "cost_price#1.currency",
        value: "EUR",
        appliesTo: ["ALL"]
    },
    {
        headerName: "item_package_dimensions#1.length.unit",
        value: "Zentimeter",
        appliesTo: ["ALL"]
    },
    {
        headerName: "item_package_dimensions#1.width.unit",
        value: "Zentimeter",
        appliesTo: ["ALL"]
    },
    {
        headerName: "item_package_dimensions#1.height.unit",
        value: "Zentimeter",
        appliesTo: ["ALL"]
    },
    {
        headerName: "item_package_weight#1.unit",
        value: "Gramm",
        appliesTo: ["ALL"]
    },
    {
        headerName: "is_assembly_required#1.value",
        value: "Nein",
        appliesTo: ["TOY_FIGURE", "TABLETOP_GAME", "TOYS_AND_GAMES", "HOBBIES", "MINIATURE_TOY_FURNISHIN", "TOY_BUILDING_BLOCK", "MINIATURE_TOY_BUILDING", "TOY_GUN"]
    },
    {
        headerName: "unit_count#1.value",
        value: 1,
        appliesTo: ["TOY_FIGURE", "FIGURINE", "MODEL", "STORAGE_BOX", "TABLETOP_GAME", "NON_RIDING_TOY_VEHICLE", "DOLL_CLOTHING", "BACKPACK", "BED_LINEN", "PUZZLES", "PILLOW", "BLANKET", "STICKER_DECAL", "TOYS_AND_GAMES", "HOBBIES", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "DRINKING_CUP", "DISPLAY_ALBUM", "DISHWARE_BOWL", "TOY_BUILDING_BLOCK", "TOY_GUN"]
    },
    {
        headerName: "unit_count#1.type.value",
        value: "stück",
        appliesTo: ["TOY_FIGURE", "FIGURINE", "MODEL", "STORAGE_BOX", "TABLETOP_GAME", "NON_RIDING_TOY_VEHICLE", "DOLL_CLOTHING", "BACKPACK", "BED_LINEN", "PUZZLES", "PILLOW", "BLANKET", "STICKER_DECAL", "TOYS_AND_GAMES", "HOBBIES", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "DRINKING_CUP", "DISPLAY_ALBUM", "DISHWARE_BOWL", "TOY_BUILDING_BLOCK", "TOY_GUN"]
    },
    {
        headerName: "street_date#1.value",
        value: (ctx: Context, row: number) => {
            const today = new Date();
            return "'" + today.toISOString().split("T")[0];
        },
        appliesTo: ["TOY_FIGURE", "APPAREL_PIN", "MOUSE_PAD", "TABLETOP_GAME", "NON_RIDING_TOY_VEHICLE", "DOLL_CLOTHING", "DESK_PAD", "PUZZLES", "TOYS_AND_GAMES", "HOBBIES", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "TOY_BUILDING_BLOCK", "TOY_GUN"]
    },
    {
        headerName: "manufacturer_minimum_age#1.value",
        value: 168,
        appliesTo: ["TOY_FIGURE", "MODEL", "TABLETOP_GAME", "NON_RIDING_TOY_VEHICLE", "DOLL_CLOTHING", "PUZZLES", "TOYS_AND_GAMES", "HOBBIES", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "TOY_BUILDING_BLOCK", "TOY_GUN"]
    },
    {
        headerName: "manufacturer_maximum_age#1.value",
        value: 1200,
        appliesTo: ["TOY_FIGURE", "TOYS_AND_GAMES", "TOY_BUILDING_BLOCK", "TOY_GUN"]
    },
    {
        headerName: "item_length_width_height#1.height.unit",
        value: "Zentimeter",
        appliesTo: ["TOY_FIGURE", "MONEY_BANK", "FIGURINE", "STORAGE_BOX", "TOTE_BAG", "DISPLAY_ALBUM", "MINIATURE_TOY_BUILDING", "TOY_GUN"]
    },
    {
        headerName: "item_length_width_height#1.length.unit",
        value: "Zentimeter",
        appliesTo: ["TOY_FIGURE", "MONEY_BANK", "FIGURINE", "STORAGE_BOX", "TOTE_BAG", "DISPLAY_ALBUM", "MINIATURE_TOY_BUILDING", "TOY_GUN"]
    },
    {
        headerName: "item_length_width_height#1.width.unit",
        value: "Zentimeter",
        appliesTo: ["TOY_FIGURE", "MONEY_BANK", "FIGURINE", "STORAGE_BOX", "TOTE_BAG", "DISPLAY_ALBUM", "MINIATURE_TOY_BUILDING", "TOY_GUN"]
    },
    {
        headerName: "warranty_description#1.value",
        value: "Keine",
        appliesTo: ["TOY_FIGURE", "MOUSE_PAD", "NON_RIDING_TOY_VEHICLE", "DOLL_CLOTHING", "DESK_PAD", "PUZZLES", "TOYS_AND_GAMES", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "TOY_BUILDING_BLOCK", "MINIATURE_TOY_BUILDING", "TOY_GUN"]
    },
    {
        headerName: "rtip_safety_warning#1.value",
        value: "Keine",
        appliesTo: ["TOY_FIGURE", "MODEL", "TABLETOP_GAME", "NON_RIDING_TOY_VEHICLE", "DOLL_CLOTHING", "PUZZLES", "STICKER_DECAL", "TOYS_AND_GAMES", "HOBBIES", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "TOY_BUILDING_BLOCK", "MINIATURE_TOY_BUILDING", "TOY_GUN"]
    },
    {
        headerName: "batteries_required#1.value",
        value: "Nein",
        appliesTo: ["TOY_FIGURE", "MONEY_BANK", "FIGURINE", "APPAREL_PIN", "STORAGE_BOX", "KEYCHAIN", "MOUSE_PAD", "TABLETOP_GAME", "NON_RIDING_TOY_VEHICLE", "TOTE_BAG", "CARRIER_BAG_CASE", "DOLL_CLOTHING", "BACKPACK", "DESK_PAD", "HOME", "DECORATIVE_MAGNET", "PILLOW", "BLANKET", "STICKER_DECAL", "TOYS_AND_GAMES", "LAMP", "THERMOS", "HOBBIES", "ADVENT_CALENDAR", "DRESS", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "DRINKING_CUP", "DISPLAY_ALBUM", "DISHWARE_BOWL", "TOY_BUILDING_BLOCK", "WALL_ART", "MINIATURE_TOY_BUILDING", "TOY_GUN"]
    },
    {
        headerName: "eu_toys_safety_directive_language#1.value",
        value: "Englisch",
        appliesTo: ["TOY_FIGURE", "MODEL", "TABLETOP_GAME", "NON_RIDING_TOY_VEHICLE", "DOLL_CLOTHING", "PUZZLES", "TOYS_AND_GAMES", "HOBBIES", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "TOY_BUILDING_BLOCK", "MINIATURE_TOY_BUILDING", "TOY_GUN"]
    },
    {
        headerName: "eu_toys_safety_directive_language#2.value",
        value: "Deutsch",
        appliesTo: ["TOY_FIGURE", "MODEL", "TABLETOP_GAME", "NON_RIDING_TOY_VEHICLE", "DOLL_CLOTHING", "PUZZLES", "TOYS_AND_GAMES", "HOBBIES", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "TOY_BUILDING_BLOCK", "MINIATURE_TOY_BUILDING", "TOY_GUN"]
    },
    {
        headerName: "eu_toys_safety_directive_age_warning#1.value",
        value: "Nicht geeignet für Kinder unter 14 Jahren",
        appliesTo: ["TOY_FIGURE", "MODEL", "TABLETOP_GAME", "NON_RIDING_TOY_VEHICLE", "DOLL_CLOTHING", "PUZZLES", "TOYS_AND_GAMES", "HOBBIES", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "TOY_BUILDING_BLOCK", "MINIATURE_TOY_BUILDING", "TOY_GUN"]
    },
    {
        headerName: "eu_toys_safety_directive_warning#1.value",
        value: "Kein Warnhinweis zutreffend",
        appliesTo: ["TOY_FIGURE", "MODEL", "TABLETOP_GAME", "NON_RIDING_TOY_VEHICLE", "DOLL_CLOTHING", "PUZZLES", "TOYS_AND_GAMES", "HOBBIES", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "TOY_BUILDING_BLOCK", "MINIATURE_TOY_BUILDING", "TOY_GUN"]
    },
    {
        headerName: "number_of_items#1.value",
        value: 1,
        appliesTo: ["MONEY_BANK", "FIGURINE", "MOUSE_PAD", "HOME", "DECORATIVE_MAGNET", "BED_LINEN", "TOWEL", "PILLOW", "BLANKET", "STICKER_DECAL", "LAMP", "MINIATURE_TOY_BUILDING"]
    },
    {
        headerName: "color#1.standardized_values#1",
        value: "Mehrfarbig",
        appliesTo: ["APPAREL_PIN", "KEYCHAIN", "TABLETOP_GAME", "TOTE_BAG", "CARRIER_BAG_CASE", "BACKPACK", "THERMOS", "DRESS", "DISPLAY_ALBUM"]
    },
    {
        headerName: "color#1.value",
        value: "Mehrfarbig",
        appliesTo: ["MONEY_BANK", "FIGURINE", "NON_RIDING_TOY_VEHICLE", "HOME", "BED_LINEN", "TOWEL", "PILLOW", "BLANKET", "STICKER_DECAL", "TOYS_AND_GAMES", "LAMP", "ART_CRAFT_KIT", "ADVENT_CALENDAR", "MINIATURE_TOY_FURNISHIN", "DOLL_CLOTHING", "WALL_ART", "TOY_FIGURE"]
    },
    {
        headerName: "is_fragile#1.value",
        value: "Nein",
        appliesTo: ["MONEY_BANK", "FIGURINE", "HOME", "DECORATIVE_MAGNET", "BED_LINEN", "BLANKET", "LAMP"]
    },
    {
        headerName: "supplier_declared_dg_hz_regulation#1.value",
        value: "Nicht zutreffend",
        appliesTo: ["MONEY_BANK", "FIGURINE", "APPAREL_PIN", "MODEL", "STORAGE_BOX", "KEYCHAIN", "MOUSE_PAD", "TABLETOP_GAME", "NON_RIDING_TOY_VEHICLE", "TOTE_BAG", "CARRIER_BAG_CASE", "DOLL_CLOTHING", "BACKPACK", "DESK_PAD", "HOME", "DECORATIVE_MAGNET", "BED_LINEN", "BLANKET", "STICKER_DECAL", "TOYS_AND_GAMES", "LAMP", "THERMOS", "PILLOW", "HOBBIES", "ADVENT_CALENDAR", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "DRINKING_CUP", "DISPLAY_ALBUM", "DISHWARE_BOWL", "TOY_BUILDING_BLOCK", "MINIATURE_TOY_BUILDING", "TOY_GUN"]
    },
    {
        headerName: "department#1.value",
        value: "Unisex",
        appliesTo: ["APPAREL_PIN", "KEYCHAIN", "TOTE_BAG", "CARRIER_BAG_CASE", "BACKPACK", "DRESS"]
    },
    {
        headerName: "metal_type#1.value",
        value: "Eisen",
        appliesTo: ["APPAREL_PIN"]
    },
    {
        headerName: "metals#1.id",
        value: 1,
        appliesTo: ["APPAREL_PIN"]
    },
    {
        headerName: "metals#1.metal_stamp.value",
        value: "Kein Metallstempel",
        appliesTo: ["APPAREL_PIN"]
    },
    {
        headerName: "metals#1.metal_type.value",
        value: "Eisen",
        appliesTo: ["APPAREL_PIN"]
    },
    {
        headerName: "metals#1.metal_weight.value",
        value: 0,
        appliesTo: ["APPAREL_PIN"]
    },
    {
        headerName: "metals#1.metal_weight.unit",
        value: "Gramm",
        appliesTo: ["APPAREL_PIN"]
    }, {
        headerName: "stones#1.id",
        value: 1,
        appliesTo: ["APPAREL_PIN"]
    },
    {
        headerName: "stones#1.type.value",
        value: "Kein Edelstein",
        appliesTo: ["APPAREL_PIN"]
    },
    {
        headerName: "stones#1.creation_method.value",
        value: "Unbekannt",
        appliesTo: ["APPAREL_PIN"]
    },
    {
        headerName: "stones#1.treatment_method.value",
        value: "Nicht behandelt",
        appliesTo: ["APPAREL_PIN"]
    }, {
        headerName: "stones#1.clarity.value",
        value: "Lupenrein",
        appliesTo: ["APPAREL_PIN"]
    },
    {
        headerName: "stones#1.color.value",
        value: "Löschen",
        appliesTo: ["APPAREL_PIN"]
    },
    {
        headerName: "stones#1.cut.value",
        value: "Exzellent",
        appliesTo: ["APPAREL_PIN"]
    },
    {
        headerName: "stones#1.shape.value",
        value: "Andere Form",
        appliesTo: ["APPAREL_PIN"]
    },
    {
        headerName: "lifecycle_supply_type#1.value",
        value: "Ganzjährig",
        appliesTo: ["APPAREL_PIN", "KEYCHAIN", "TOTE_BAG", "CARRIER_BAG_CASE", "BACKPACK", "DRESS"]
    },
    {
        headerName: "target_audience_keyword#1.value",
        value: "Unisex – Erwachsene",
        appliesTo: ["TOY_FIGURE", "MODEL", "TABLETOP_GAME", "DOLL_CLOTHING", "PUZZLES", "TOYS_AND_GAMES", "HOBBIES", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "TOY_BUILDING_BLOCK", "TOY_GUN", "NON_RIDING_TOY_VEHICLE"]
    },
    {
        headerName: "age_range_description#1.value",
        value: "Erwachsene",
        appliesTo: ["MODEL", "TABLETOP_GAME", "NON_RIDING_TOY_VEHICLE", "TOTE_BAG", "DOLL_CLOTHING", "BACKPACK", "TOYS_AND_GAMES", "HOBBIES", "DRESS", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "TOY_BUILDING_BLOCK", "MINIATURE_TOY_BUILDING", "TOY_GUN"]
    },
    {
        headerName: "item_dimensions#1.length.unit",
        value: "Zentimeter",
        appliesTo: ["MODEL", "NON_RIDING_TOY_VEHICLE", "TOYS_AND_GAMES", "HOBBIES", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "TOY_BUILDING_BLOCK"]
    },
    {
        headerName: "item_dimensions#1.width.unit",
        value: "Zentimeter",
        appliesTo: ["MODEL", "NON_RIDING_TOY_VEHICLE", "TOYS_AND_GAMES", "HOBBIES", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "TOY_BUILDING_BLOCK"]
    },
    {
        headerName: "item_dimensions#1.height.unit",
        value: "Zentimeter",
        appliesTo: ["MODEL", "NON_RIDING_TOY_VEHICLE", "TOYS_AND_GAMES", "HOBBIES", "PRETEND_PLAY_TOY", "MINIATURE_TOY_FURNISHIN", "TOY_BUILDING_BLOCK"]
    },
    {
        headerName: "number_of_boxes#1.value",
        value: 1,
        appliesTo: ["MODEL", "STORAGE_BOX", "MOUSE_PAD", "HOME", "STICKER_DECAL", "TOYS_AND_GAMES", "MINIATURE_TOY_FURNISHIN", "TOY_GUN"]
    },
    {
        headerName: "target_gender#1.value",
        value: "Unisex",
        appliesTo: ["KEYCHAIN", "TOTE_BAG", "BACKPACK", "DRESS"]
    },
    {
        headerName: "weave_type#1.value",
        value: "Dobby",
        appliesTo: ["KEYCHAIN", "DRESS"]
    },
    {
        headerName: "rtip_order_aggregate_type#1.value",
        value: "Jede/-r/-s",
        appliesTo: ["KEYCHAIN", "TOTE_BAG", "CARRIER_BAG_CASE", "BACKPACK", "DRESS"]
    },
    {
        headerName: "item_thickness#1.decimal_value",
        value: 2,
        appliesTo: ["MOUSE_PAD", "DESK_PAD", "PILLOW", "BLANKET"]
    },
    {
        headerName: "item_thickness#1.unit",
        value: "Millimeter",
        appliesTo: ["MOUSE_PAD", "DESK_PAD", "PILLOW", "BLANKET"]
    },
    {
        headerName: "gdpr_risk#1.value",
        value: "Keine elektronischen Informationen gespeichert",
        appliesTo: ["MOUSE_PAD"]
    },
    {
        headerName: "language#1.type",
        value: "Unbekannt",
        appliesTo: ["TABLETOP_GAME"]
    },
    {
        headerName: "language#1.value",
        value: "Keine",
        appliesTo: ["TABLETOP_GAME"]
    },
    {
        headerName: "power_plug_type#1.value",
        value: "Kein Stecker",
        appliesTo: ["TABLETOP_GAME", "HOME", "TOYS_AND_GAMES", "LAMP"]
    },
    {
        headerName: "style#1.value",
        value: "Zeitgenössisch",
        appliesTo: ["TOTE_BAG", "CARRIER_BAG_CASE", "BACKPACK", "DRESS"]
    },
    {
        headerName: "strap_type#1.value",
        value: "Schulter",
        appliesTo: ["TOTE_BAG"]
    },
    {
        headerName: "strap_type#1.value",
        value: "Gepolsterter Gurt",
        appliesTo: ["BACKPACK"]
    },
    {
        headerName: "number_of_compartments#1.value",
        value: 1,
        appliesTo: ["TOTE_BAG", "BACKPACK"]
    },
    {
        headerName: "number_of_pockets#1.value",
        value: 1,
        appliesTo: ["TOTE_BAG"]
    },
    {
        headerName: "closure#1.type#1.value",
        value: "Reißverschluss",
        appliesTo: ["TOTE_BAG", "BACKPACK", "DRESS"]
    },
    {
        headerName: "item_length_width#1.length.unit",
        value: "Zentimeter",
        appliesTo: ["DOLL_CLOTHING", "TOWEL", "PUZZLES", "PILLOW", "PILLOW", "BLANKET", "WALL_ART"]
    },
    {
        headerName: "item_length_width#1.width.unit",
        value: "Zentimeter",
        appliesTo: ["DOLL_CLOTHING", "TOWEL", "PUZZLES", "PILLOW", "PILLOW", "BLANKET", "WALL_ART"]
    },
    {
        headerName: "water_resistance_level#1.value",
        value: "Feuchtigkeitsresistent",
        appliesTo: ["BACKPACK", "DRESS"]
    },
    {
        headerName: "storage_volume#1.unit",
        value: "Liter",
        appliesTo: ["BACKPACK"]
    },
    {
        headerName: "item_depth_width_height#1.depth.unit",
        value: "Zentimeter",
        appliesTo: ["BACKPACK", "LAMP"]
    },
    {
        headerName: "item_depth_width_height#1.height.unit",
        value: "Zentimeter",
        appliesTo: ["BACKPACK", "LAMP"]
    },
    {
        headerName: "item_depth_width_height#1.width.unit",
        value: "Zentimeter",
        appliesTo: ["BACKPACK", "LAMP"]
    },
    {
        headerName: "pattern#1.value",
        value: "Bedruckt",
        appliesTo: ["BED_LINEN"]
    },
    {
        headerName: "contains_food_or_beverage#1.value",
        value: "Nein",
        appliesTo: ["TOYS_AND_GAMES"]
    },
    {
        headerName: "accepted_voltage_frequency#1.value",
        value: "220 bis 240 Volt und 50 Hertz",
        appliesTo: ["TOYS_AND_GAMES", "LAMP"]
    },
    {
        headerName: "power_source_type#1.value",
        value: "Kabelgebunden",
        appliesTo: ["LAMP"]
    },
    {
        headerName: "eu_energy_label_efficiency_class#1.value",
        value: "A+++ to G",
        appliesTo: ["LAMP"]
    },
    {
        headerName: "efficiency#1.value",
        value: "G",
        appliesTo: ["LAMP"]
    },
    {
        headerName: "eu_energy_efficiency_rating#1.value",
        value: "Nicht bewertet",
        appliesTo: ["LAMP"]
    },
    {
        headerName: "eprel_registration_number#1.value",
        value: "keine",
        appliesTo: ["LAMP"]
    },
    {
        headerName: "capacity#1.unit",
        value: "Milliliter",
        appliesTo: ["THERMOS", "DRINKING_CUP"]
    },
    {
        headerName: "item_width_height#1.height.unit",
        value: "Zentimeter",
        appliesTo: ["THERMOS", "DRINKING_CUP", "DISHWARE_BOWL"]
    },
    {
        headerName: "item_width_height#1.width.unit",
        value: "Zentimeter",
        appliesTo: ["THERMOS", "DRINKING_CUP", "DISHWARE_BOWL"]
    },
    {
        headerName: "apparel_size#1.size_system",
        value: "DE/NL/SE/PL",
        appliesTo: ["DRESS"]
    },
    {
        headerName: "apparel_size#1.size_class",
        value: "Alphanumerisch",
        appliesTo: ["DRESS"]
    },
    {
        headerName: "apparel_size#1.size",
        value: "1",
        appliesTo: ["DRESS"]
    },
    {
        headerName: "apparel_size#1.body_type",
        value: "Regular",
        appliesTo: ["DRESS"]
    },
    {
        headerName: "apparel_size#1.height_type",
        value: "Kurz",
        appliesTo: ["DRESS"]
    },
    {
        headerName: "item_length_description#1.value",
        value: "Teelang",
        appliesTo: ["DRESS"]
    },
    {
        headerName: "care_instructions#1.value",
        value: "Nur Handwäsche",
        appliesTo: ["DRESS"]
    },
    {
        headerName: "neck#1.neck_style#1.value",
        value: "Crew Neck",
        appliesTo: ["DRESS"]
    },
    {
        headerName: "neck#1.neck_style#1.value",
        value: "Fledermausärmel",
        appliesTo: ["DRESS"]
    },
    {
        headerName: "is_refurbished#1.value",
        value: "Nein",
        appliesTo: ["DRINKING_CUP", "DISHWARE_BOWL"]
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
