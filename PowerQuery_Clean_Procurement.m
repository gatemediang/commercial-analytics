let
    // Replace the path below with the location of 01_Procurement_Spend.csv.
    Source = Csv.Document(
        File.Contents("C:\\Portfolio\\01_Procurement_Spend.csv"),
        [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    TrimmedText = Table.TransformColumns(
        PromotedHeaders,
        {
            {"Supplier", Text.Trim, type text},
            {"Category", Text.Trim, type text},
            {"SiteID", Text.Trim, type text},
            {"InvoiceID", Text.Trim, type text}
        }
    ),
    StandardisedCategory = Table.TransformColumns(
        TrimmedText,
        {{"Category", each Text.Proper(_), type text}}
    ),
    TypedColumns = Table.TransformColumnTypes(
        StandardisedCategory,
        {
            {"InvoiceDate", type date},
            {"Quantity", Int64.Type},
            {"UnitCost", type number},
            {"GrossSpend", type number},
            {"PaymentTermsDays", Int64.Type}
        }
    ),
    DuplicateCount = Table.Group(
        TypedColumns,
        {"InvoiceID"},
        {{"InvoiceCount", each Table.RowCount(_), Int64.Type}}
    ),
    JoinedDuplicateCount = Table.NestedJoin(
        TypedColumns,
        {"InvoiceID"},
        DuplicateCount,
        {"InvoiceID"},
        "DQ",
        JoinKind.LeftOuter
    ),
    ExpandedDQ = Table.ExpandTableColumn(
        JoinedDuplicateCount,
        "DQ",
        {"InvoiceCount"},
        {"InvoiceCount"}
    ),
    QualityFlag = Table.AddColumn(
        ExpandedDQ,
        "DataQualityFlag",
        each if [InvoiceID] = null or [InvoiceID] = "" then "Missing invoice"
             else if [Supplier] = null or [Supplier] = "" then "Missing supplier"
             else if [SiteID] = null or [SiteID] = "" then "Missing site"
             else if [GrossSpend] = null then "Missing spend"
             else if [InvoiceCount] > 1 then "Duplicate invoice"
             else "OK",
        type text
    )
in
    QualityFlag
