import argparse
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser(description='Format Word')
    parser.add_argument('--self-test', metavar='REPORT_JSON', type=Path,
                        help='Verifica o executável com arquivos temporários e grava o resultado.')
    args = parser.parse_args()
    if args.self_test:
        from app.selftest import run_self_test
        raise SystemExit(run_self_test(args.self_test))
    from app.ui import FormatWordApp
    app = FormatWordApp()
    app.mainloop()


if __name__ == '__main__':
    main()
