// Power Query (M): clean and flag procurement spend, mirroring Cleaned_Procurement.
// v2 fixes: Text.Proper turned "IT & Telecoms" into "It & Telecoms", so categories and
// suppliers are now standardised by a case-insensitive join to master lists instead.
// Setup: load tblSuppliers and tblTargets from the workbook as queries named
// SupplierMaster and CategoryMaster, then paste this into a Blank Query.
let
    Source = Csv.Document(
        File.Contents("C:\Portfolio\01_Procurement_Spend.csv"),   // change to your path / SharePoint
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed = Table.TransformColumnTypes(Promoted, {
        {"InvoiceDate", type date}, {"Quantity", Int64.Type}, {"UnitCost", type number},
        {"GrossSpend", type number}, {"PaymentTermsDays", Int64.Type}}, "en-GB"),
    Trimmed = Table.TransformColumns(Typed, {
        {"Supplier", Text.Trim, type text}, {"Category", Text.Trim, type text},
        {"SiteID", Text.Trim, type text}, {"InvoiceID", Text.Trim, type text}}),

    // case-insensitive keys for joining to the masters
    WithKeys = Table.AddColumn(
        Table.AddColumn(Trimmed, "SupKey", each Text.Lower([Supplier] ?? ""), type text),
        "CatKey", each Text.Lower([Category] ?? ""), type text),
    SupMaster = Table.AddColumn(SupplierMaster, "SupKey", each Text.Lower([Supplier]), type text),
    CatMaster = Table.AddColumn(CategoryMaster, "CatKey", each Text.Lower([Category]), type text),

    JoinSup = Table.NestedJoin(WithKeys, {"SupKey"}, SupMaster, {"SupKey"}, "Sup", JoinKind.LeftOuter),
    ExpandSup = Table.ExpandTableColumn(JoinSup, "Sup", {"Supplier", "Category"}, {"SupplierClean", "MasterCategory"}),
    JoinCat = Table.NestedJoin(ExpandSup, {"CatKey"}, CatMaster, {"CatKey"}, "Cat", JoinKind.LeftOuter),
    ExpandCat = Table.ExpandTableColumn(JoinCat, "Cat", {"Category", "TargetSavingsPct"}, {"CategoryClean", "SavingsTarget"}),

    // duplicate tests: same ID anywhere, and same ID + supplier + amount (true duplicate)
    IdCounts = Table.Group(ExpandCat, {"InvoiceID"}, {{"IdCount", each Table.RowCount(_), Int64.Type}}),
    WithIdCount = Table.ExpandTableColumn(
        Table.NestedJoin(ExpandCat, {"InvoiceID"}, IdCounts, {"InvoiceID"}, "Ids", JoinKind.LeftOuter),
        "Ids", {"IdCount"}),
    Indexed = Table.AddIndexColumn(WithIdCount, "RowNo", 1, 1, Int64.Type),
    FullKey = Table.AddColumn(Indexed, "FullKey",
        each [InvoiceID] & "|" & ([SupplierClean] ?? "") & "|" & Number.ToText([GrossSpend] ?? 0, "F2"), type text),
    FirstSeen = Table.Group(FullKey, {"FullKey"}, {{"FirstRow", each List.Min([RowNo]), Int64.Type}}),
    WithFirst = Table.ExpandTableColumn(
        Table.NestedJoin(FullKey, {"FullKey"}, FirstSeen, {"FullKey"}, "F", JoinKind.LeftOuter), "F", {"FirstRow"}),

    Flagged = Table.AddColumn(WithFirst, "DataQualityFlag", each
        if [Supplier] = null or [Supplier] = "" then "Missing supplier"
        else if [SiteID] = null or [SiteID] = "" then "Missing site"
        else if [GrossSpend] = null then "Missing spend"
        else if [RowNo] > [FirstRow] then "Duplicate invoice"
        else if [IdCount] > 1 then "Duplicate ID - review"
        else if [MasterCategory] <> null and Text.Lower([MasterCategory]) <> [CatKey] then "Category mismatch"
        else "OK", type text),
    NetSpend = Table.AddColumn(Flagged, "NetSpend_Clean", each
        if List.Contains({"Missing spend", "Duplicate invoice"}, [DataQualityFlag]) then 0 else [GrossSpend], type number),
    Potential = Table.AddColumn(NetSpend, "Potential_Saving", each [NetSpend_Clean] * ([SavingsTarget] ?? 0), type number),
    Sorted = Table.Sort(Potential, {{"RowNo", Order.Ascending}}),
    Result = Table.SelectColumns(Sorted, {
        "InvoiceDate", "SiteID", "SiteName", "Region", "SupplierClean", "CategoryClean", "InvoiceID",
        "GrossSpend", "Contracted", "DataQualityFlag", "NetSpend_Clean", "SavingsTarget", "Potential_Saving"})
in
    Result
