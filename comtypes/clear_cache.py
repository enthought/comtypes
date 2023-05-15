import argparse
import os
import sys
from shutil import rmtree  # TESTS ASSUME USE OF RMTREE


def main():
    parser = argparse.ArgumentParser(
        prog="py -m comtypes.clear_cache", description="Removes comtypes cache folders."
    )
    parser.add_argument(
        "-y", help="Pre-approve deleting all folders", action="store_true"
    )
    args = parser.parse_args()

    if not args.y:
        confirm = input("Remove comtypes cache directories? (y/n): ")
        if confirm.lower() != "y":
            print("Cache directories NOT removed")
            return

    # change cwd to avoid import from local folder during installation process
    work_dir = os.getcwd()
    try:
        os.chdir(os.path.dirname(sys.executable))
        import comtypes.client
    except ImportError:
        print("Could not import comtypes", file=sys.stderr)
        sys.exit(1)
    os.chdir(work_dir)

    # there are two possible locations for the cache folder (in the comtypes
    # folder in site-packages if that is writable, otherwise in APPDATA)
    # fortunately, by deleting the first location returned by _find_gen_dir()
    # we make it un-writable, so calling it again gives us the APPDATA location
    for _ in range(2):
        dir_path = comtypes.client._find_gen_dir()
        rmtree(dir_path)
        print(f"Removed directory \"{dir_path}\"")


if __name__ == "__main__":
    main()
