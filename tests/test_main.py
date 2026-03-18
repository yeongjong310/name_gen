from name_gen.main import main


def test_main(capsys):
    main()
    captured = capsys.readouterr()
    assert captured.out == "Hello from name_gen!\n"
