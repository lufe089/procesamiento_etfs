"""
Tests for esg_processor module.
"""

import os
import tempfile

import pandas as pd
import pytest

from src.esg_processor import (
    build_esg_base,
    compute_simple_average,
    compute_weighted_average,
    load_data,
    save_results,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def sample_df():
    """DataFrame mínimo con dos ETFs y varias empresas."""
    return pd.DataFrame(
        {
            "ETF": ["ETF_A", "ETF_A", "ETF_A", "ETF_B", "ETF_B"],
            "Company": ["Empresa1", "Empresa2", "Empresa3", "Empresa4", "Empresa5"],
            "ESG_Score": [80.0, 60.0, 40.0, 70.0, 30.0],
            "Market_Cap": [100.0, 200.0, 300.0, 500.0, 500.0],
        }
    )


@pytest.fixture
def single_etf_df():
    """DataFrame con un único ETF."""
    return pd.DataFrame(
        {
            "ETF": ["ETF_X", "ETF_X"],
            "Company": ["CompA", "CompB"],
            "ESG_Score": [50.0, 90.0],
            "Market_Cap": [400.0, 100.0],
        }
    )


# ---------------------------------------------------------------------------
# load_data
# ---------------------------------------------------------------------------

def test_load_data_reads_csv_correctly(tmp_path):
    csv_content = "ETF,Company,ESG_Score,Market_Cap\nSPY,Apple,72.5,2800000\nSPY,Microsoft,85.3,2500000\n"
    csv_file = tmp_path / "holdings.csv"
    csv_file.write_text(csv_content)

    df = load_data(str(csv_file))

    assert list(df.columns) == ["ETF", "Company", "ESG_Score", "Market_Cap"]
    assert len(df) == 2
    assert pd.api.types.is_numeric_dtype(df["ESG_Score"])
    assert pd.api.types.is_numeric_dtype(df["Market_Cap"])


def test_load_data_raises_on_missing_columns(tmp_path):
    csv_content = "ETF,Company,ESG_Score\nSPY,Apple,72.5\n"
    csv_file = tmp_path / "bad.csv"
    csv_file.write_text(csv_content)

    with pytest.raises(ValueError, match="Market_Cap"):
        load_data(str(csv_file))


# ---------------------------------------------------------------------------
# compute_simple_average
# ---------------------------------------------------------------------------

def test_simple_average_values(sample_df):
    result = compute_simple_average(sample_df)

    assert set(result.columns) == {"ETF", "ESG_Promedio_Simple"}
    etf_a = result.loc[result["ETF"] == "ETF_A", "ESG_Promedio_Simple"].iloc[0]
    etf_b = result.loc[result["ETF"] == "ETF_B", "ESG_Promedio_Simple"].iloc[0]

    # ETF_A: (80 + 60 + 40) / 3 = 60.0
    assert etf_a == pytest.approx(60.0, rel=1e-4)
    # ETF_B: (70 + 30) / 2 = 50.0
    assert etf_b == pytest.approx(50.0, rel=1e-4)


def test_simple_average_single_etf(single_etf_df):
    result = compute_simple_average(single_etf_df)
    assert len(result) == 1
    assert result.iloc[0]["ESG_Promedio_Simple"] == pytest.approx(70.0, rel=1e-4)


def test_simple_average_returns_one_row_per_etf(sample_df):
    result = compute_simple_average(sample_df)
    assert len(result) == 2


# ---------------------------------------------------------------------------
# compute_weighted_average
# ---------------------------------------------------------------------------

def test_weighted_average_values(sample_df):
    result = compute_weighted_average(sample_df)

    assert set(result.columns) == {"ETF", "ESG_Promedio_Ponderado"}
    etf_a = result.loc[result["ETF"] == "ETF_A", "ESG_Promedio_Ponderado"].iloc[0]
    etf_b = result.loc[result["ETF"] == "ETF_B", "ESG_Promedio_Ponderado"].iloc[0]

    # ETF_A: (80*100 + 60*200 + 40*300) / (100+200+300) = (8000+12000+12000)/600 = 32000/600 ≈ 53.3333
    assert etf_a == pytest.approx(32000 / 600, rel=1e-4)
    # ETF_B: (70*500 + 30*500) / (500+500) = (35000+15000)/1000 = 50.0
    assert etf_b == pytest.approx(50.0, rel=1e-4)


def test_weighted_average_single_etf(single_etf_df):
    result = compute_weighted_average(single_etf_df)
    # ETF_X: (50*400 + 90*100) / (400+100) = (20000+9000)/500 = 58.0
    assert result.iloc[0]["ESG_Promedio_Ponderado"] == pytest.approx(58.0, rel=1e-4)


def test_weighted_average_returns_one_row_per_etf(sample_df):
    result = compute_weighted_average(sample_df)
    assert len(result) == 2


def test_weighted_average_zero_market_cap():
    df = pd.DataFrame(
        {
            "ETF": ["ETF_Z", "ETF_Z"],
            "Company": ["CompA", "CompB"],
            "ESG_Score": [60.0, 80.0],
            "Market_Cap": [0.0, 0.0],
        }
    )
    result = compute_weighted_average(df)
    assert pd.isna(result.iloc[0]["ESG_Promedio_Ponderado"])


# ---------------------------------------------------------------------------
# build_esg_base
# ---------------------------------------------------------------------------

def test_build_esg_base_columns(sample_df):
    base = build_esg_base(sample_df)
    assert set(base.columns) == {"ETF", "ESG_Promedio_Simple", "ESG_Promedio_Ponderado"}


def test_build_esg_base_row_count(sample_df):
    base = build_esg_base(sample_df)
    assert len(base) == 2


def test_build_esg_base_etf_a_values(sample_df):
    base = build_esg_base(sample_df)
    row = base.loc[base["ETF"] == "ETF_A"].iloc[0]
    assert row["ESG_Promedio_Simple"] == pytest.approx(60.0, rel=1e-4)
    assert row["ESG_Promedio_Ponderado"] == pytest.approx(32000 / 600, rel=1e-4)


# ---------------------------------------------------------------------------
# save_results
# ---------------------------------------------------------------------------

def test_save_results_creates_csv(sample_df, tmp_path):
    base = build_esg_base(sample_df)
    output_file = str(tmp_path / "output.csv")
    save_results(base, output_file)

    assert os.path.exists(output_file)
    loaded = pd.read_csv(output_file)
    assert set(loaded.columns) == {"ETF", "ESG_Promedio_Simple", "ESG_Promedio_Ponderado"}
    assert len(loaded) == 2
