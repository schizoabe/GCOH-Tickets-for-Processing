from collections import defaultdict
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "HierarchyFiles" / "parentnode.xlsx"
OUTPUT_FILE = BASE_DIR / "HierarchyFiles" / "hierarchy_output.xlsx"

EDGE_SHEET = "ParentChild"
NAME_SHEET = "ObjectNames"


def read_edges(path):
    """Read Parent/Child pairs from the SAP download, skipping the '*' metadata header."""
    raw = pd.read_excel(path, sheet_name=EDGE_SHEET, header=None, dtype=str)
    raw = raw.iloc[:, :2]
    raw.columns = ["Parent", "Node"]
    raw = raw.apply(lambda col: col.str.strip())
    # Metadata lines start with '*'; the edge list begins after the last one ("*CCTRG")
    header_rows = raw.index[raw["Parent"].str.startswith("*", na=False)]
    if len(header_rows):
        raw = raw.loc[header_rows[-1] + 1:]
    edges = raw.dropna(how="any")
    return edges.drop_duplicates().reset_index(drop=True)


def read_names(path):
    names = pd.read_excel(path, sheet_name=NAME_SHEET, dtype=str)
    names = names.dropna(subset=["Node"])
    return dict(zip(names["Node"].str.strip(), names["Name"].fillna("").str.strip()))


def build_paths(edges):
    """Return the path from its root to every node, in depth-first (file) order."""
    children = defaultdict(list)
    for parent, node in edges.itertuples(index=False):
        children[parent].append(node)

    child_set = set(edges["Node"])
    roots = list(dict.fromkeys(p for p in edges["Parent"] if p not in child_set))

    paths = []
    stack = [[r] for r in reversed(roots)]
    while stack:
        path = stack.pop()
        paths.append(path)
        for child in reversed(children.get(path[-1], [])):
            if child in path:
                print(f"WARNING: cycle skipped at {path[-1]} -> {child}")
                continue
            stack.append(path + [child])

    reached = {n for p in paths for n in p}
    unreached = (set(children) | child_set) - reached
    if unreached:
        print(f"WARNING: {len(unreached)} nodes only reachable through cycles, "
              f"e.g. {sorted(unreached)[:5]}")
    return paths


def hiymaker(input_file=INPUT_FILE, output_file=OUTPUT_FILE):
    print("running hierarchy maker")

    edges = read_edges(input_file)
    names = read_names(input_file)
    if edges.empty:
        raise ValueError(f"No Parent/Child rows found in '{EDGE_SHEET}' of {input_file}")

    paths = build_paths(edges)

    def label(node):
        name = names.get(node)
        return f"{node} - [{name}]" if name else node

    missing = {n for p in paths for n in p if n not in names}
    if missing:
        print(f"WARNING: {len(missing)} nodes have no name in '{NAME_SHEET}'")

    depth = max(len(p) for p in paths)
    output_df = pd.DataFrame(
        [[label(n) for n in p] + [None] * (depth - len(p)) for p in paths],
        columns=[f"level_{i}" for i in range(depth)],
    )

    output_df.to_excel(output_file, sheet_name="Output", index=False)

    print(f"hierarchy generation complete: {len(output_df)} rows, "
          f"{depth} levels -> {output_file}")
    return output_file


if __name__ == "__main__":
    hiymaker()
