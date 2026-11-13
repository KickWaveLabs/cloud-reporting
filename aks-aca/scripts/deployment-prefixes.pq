// =====================================================================================================
// Power Query: PROD deployments whose name does not start with an application prefix
// =====================================================================================================
// Where to run: a blank Excel workbook > Data > Get Data > From Other Sources > Blank Query >
// Advanced Editor. Paste this, set FilePath to the platform team's reportDeployment export, and Close & Load.
//
// - The result groups the deployments by cluster, namespace and prefix (first word of the name), with up
//   to three example names. Mark the ones that are not applications (tools, agents, gateways) and send
//   the list back; it feeds the cluster category (application or infrastructure).
// - For the full list instead of the summary, change the last line to `in WithType`.
// - Add a prefix to AppPrefixes when it turns out to be an application (for example "async").
// - For a CSV export, replace the Book and Sheet steps with:
//     Sheet = Table.PromoteHeaders(Csv.Document(File.Contents(FilePath), [Delimiter = ";", Encoding = 65001])),
// =====================================================================================================

let
    FilePath = "C:\Users\<user>\Downloads\reportDeployment.xlsx",
    AppPrefixes = {"business", "channel", "app", "core"},

    Book = Excel.Workbook(File.Contents(FilePath), true),
    Sheet = Table.SelectRows(Book, each [Kind] = "Sheet"){0}[Data],
    Kept = Table.SelectColumns(Sheet, {"resourceName", "environment", "Namespace", "Deployment"}),
    Prod = Table.SelectRows(Kept, each Text.Upper(Text.From([environment])) = "PROD"),
    WithPrefix = Table.AddColumn(Prod, "prefix",
        each Text.Lower(Text.BeforeDelimiter(Text.From([Deployment]), "-")), type text),
    WithType = Table.AddColumn(WithPrefix, "app_prefix",
        each if List.Contains(AppPrefixes, [prefix]) then "YES" else "NO", type text),
    NotApp = Table.SelectRows(WithType, each [app_prefix] = "NO"),
    Summary = Table.Group(NotApp, {"resourceName", "Namespace", "prefix"}, {
        {"deployments", each Table.RowCount(_), Int64.Type},
        {"examples", each Text.Combine(List.FirstN([Deployment], 3), "; "), type text}}),
    Sorted = Table.Sort(Summary, {{"resourceName", Order.Ascending}, {"deployments", Order.Descending}})
in
    Sorted
