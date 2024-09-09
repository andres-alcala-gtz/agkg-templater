import click
import pathlib

import search
import update


if __name__ == "__main__":

    print()
    directory = click.prompt("Directory", type=click.Path(exists=True, file_okay=False, dir_okay=True))
    print()

    directory_src = pathlib.Path(f"{directory}")

    data = search.search(directory_src)

    identifier = data["«identifier»"]
    name = data["«name»"]

    directory_tpl = pathlib.Path(f"_Templates")
    directory_dst = pathlib.Path(f"{identifier} - {name}")

    _ = update.update(directory_tpl, directory_dst, data)

    print()
