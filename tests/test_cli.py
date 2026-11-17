from finprep.cli import main


def test_demo_writes_outputs(tmp_path):
    out, data = tmp_path / "out", tmp_path / "data"
    code = main(["demo", "-o", str(out), "--data-dir", str(data)])
    assert code == 0
    for name in (
        "tidy_quarterly.csv",
        "tidy_annual.csv",
        "summary_quarterly.csv",
        "quality_report.md",
    ):
        assert (out / name).stat().st_size > 0
    assert (data / "raw_fnltt.csv").exists()


def test_run_reproduces_demo_from_csv(tmp_path):
    out, data = tmp_path / "out", tmp_path / "data"
    main(["demo", "-o", str(out), "--data-dir", str(data)])
    code = main(
        [
            "run",
            str(data / "raw_fnltt.csv"),
            "--segments",
            str(data / "ir_segments.csv"),
            "-o",
            str(tmp_path / "rerun"),
        ]
    )
    assert code == 0
    first = (out / "tidy_quarterly.csv").read_bytes()
    assert (tmp_path / "rerun" / "tidy_quarterly.csv").read_bytes() == first


def test_fetch_without_key_fails_cleanly(tmp_path, monkeypatch, capsys):
    monkeypatch.delenv("DART_API_KEY", raising=False)
    monkeypatch.setattr("finprep.fetch.dart.CACHE_DIR", tmp_path / "cache")
    code = main(["fetch", "005930", "-o", str(tmp_path)])
    assert code == 2
    assert "DART_API_KEY" in capsys.readouterr().err
