def test_package_entry_imports():
    from tender_formatter import __version__
    from tender_formatter.main import main

    assert callable(main)
    assert __version__ == "0.1.0"
