"""Excel compare test suite."""
import io

import pandas as pd

import compare

def test_parser():
    cfg = compare.build_parser()
    opt = cfg.parse_args(["test1.xlsx", "test2.xlsx", "Sheet 1", "Col1", "Col2", "-o", "output.xlsx"])
    assert opt.path1 == "test1.xlsx"
    assert opt.path2 == "test2.xlsx"
    assert opt.output_path ==  "output.xlsx"
    assert opt.sheetname == "Sheet 1"
    assert opt.key_column == ["Col1", "Col2"]
    assert opt.skiprows is None


def build_excel_stream(df, sheetname):
    """Create an excel workbook as a file-like object."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheetname, index=False)
    return output


def sample_xlsx(df_1, df_2):
    xlsx_1 = build_excel_stream(df_1, "Sheet1")
    xlsx_2 = build_excel_stream(df_2, "Sheet1")
    return xlsx_1, xlsx_2


def sample_dfs():
    df_1 = pd.DataFrame({
        "ID": [123456, 654321, 543219, 432198, 765432],
        "Name": ["Lemonade", "Cola", "Orange", "Fruit Punch", "Tobacco"],
        "Flavour Description": ["Fuzzy", "Fuzzy", "Fuzzy", "Fuzzy", "Smoky"],
    })
    df_2 = pd.DataFrame({
        "ID": [123456, 654321, 543219, 432198, 876543],
        "Name": ["Lemonade", "Cola", "Orange", "Fruit Punch", "Soda"],
        "Flavour Description": ["Fuzzy", "Bubbly", "Fuzzy", "Fuzzy", "Sugary"],
    })
    return df_1, df_2


def run_assertion(diff):
    changed = diff["changed"]
    assert len(changed) == 1
    assert changed.iloc[0]["Flavour Description"] == "Fuzzy ---> Bubbly"
    added = diff["added"]
    assert len(added) == 1
    assert added.iloc[0]["Flavour Description"] == "Sugary"
    removed = diff["removed"]
    assert len(removed) == 1
    assert removed.iloc[0]["Flavour Description"] == "Smoky"
    print("OK.")


def test_single_index():
    df_1, df_2 = sample_dfs()
    diff = compare.diff_pd(df_1, df_2, ["ID"])
    run_assertion(diff)


def test_single_index_excel():
    xlsx_1, xlsx_2 = sample_xlsx(*sample_dfs())
    diff_io = io.BytesIO()
    compare.compare_excel(xlsx_1, xlsx_2, diff_io, "Sheet1", "ID")
    diff = pd.read_excel(diff_io, sheet_name=None)
    run_assertion(diff)


def sample_multiindex_dfs():
    df_1 = pd.DataFrame({
        "ID": [123456, 123456, 654321, 543219, 432198, 765432],
        "Name": ["Lemonade", "Lemonade", "Cola", "Orange", "Fruit Punch", "Tobacco"],
        "Flavour ID": [1, 2, None, None, None, None],
        "Flavour Description": ["Fuzzy", "Fuzzy", "Fuzzy", "Fuzzy", "Fuzzy", "Smoky"],
    })
    df_2 = pd.DataFrame({
        "ID": [123456, 123456, 654321, 543219, 432198, 876543],
        "Name": ["Lemonade", "Lemonade", "Cola", "Orange", "Fruit Punch", "Soda"],
        "Flavour ID": [1, 2, None, None, None, None],
        "Flavour Description": ["Fuzzy", "Bubbly", "Fuzzy", "Fuzzy", "Fuzzy", "Sugary"],
    })
    return df_1, df_2


def test_multiindex():
    df_1, df_2 = sample_multiindex_dfs()
    diff = compare.diff_pd(df_1, df_2, ["ID", "Flavour ID"])
    run_assertion(diff)


def test_multiindex_excel():
    xlsx_1, xlsx_2 = sample_xlsx(*sample_multiindex_dfs())
    diff_io = io.BytesIO()
    compare.compare_excel(xlsx_1, xlsx_2, diff_io, "Sheet1", ["ID", "Flavour ID"])
    diff = pd.read_excel(diff_io, sheet_name=None)
    run_assertion(diff)
    

def test_no_diffs():
    df_1, _ = sample_multiindex_dfs()
    diff = compare.diff_pd(df_1, df_1, ["ID", "Flavour ID"])
    assert not diff
    print("OK.")


if __name__ == '__main__':
    test_multiindex()
    test_multiindex_excel()
    test_single_index()
    test_single_index_excel()
    test_parser()
    test_no_diff()
