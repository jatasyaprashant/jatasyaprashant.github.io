"""Rebuild the publication map from a local v1.1 dataset checkout.

Install openpyxl and networkx, then run:
    python tools/build_tcs_page.py <dataset-directory>
"""
import argparse
import csv
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from collections import Counter, defaultdict
from html import escape
import json
import math
from pathlib import Path
import re
from string import Template
from urllib.parse import urlsplit

import networkx as nx
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
URL = "https://jatasyaprashant.github.io/tcs-research/"
TITLE = "What TCS Research publishes: a 2023–2026 publication map"
DESCRIPTION = ("An independent map of TCS Research papers, venues, research groups, "
               "and collaborations through 23 November 2026.")
CSV_FIELDS = ("title", "year", "venue", "venue_tier", "domain", "status",
              "authors", "tcs_authors", "doi", "arxiv_url")
FOCUS = (
    ("Unlearning", "Papers classified as Unlearning and titles matching related terms such as forgetting, leakage and attribution.",
     "Unlearning", r"unlearn|forget|erasure|attribution|leakage|alignment faking|realign"),
    ("SLM, efficient models and edge", "Title matches cover small models, compression, distillation, on-device computing and related terms.",
     None, r"small language|SLM|distill|quantiz|tinyml|on-device|edge|compress|neuromorph|spiking"),
    ("Vision-language and multimodal", "Title matches cover vision-language models, multimodal work and visual question answering.",
     None, r"vision-language|VLM|multimodal|VQA"),
    ("Agents", "Papers classified in the Agents domain.", "Agents", None),
)
GROUPS = {
    "unlearning": ("Shirish Subhash Karande", "Vivek Srivastava", "Ashutosh Ranjan",
                   "Ashok Urlana", "Manasi S. Patwardhan", "Snigdha Paul"),
    "edge": ("Arpan Pal", "S. Dey", "Arijit Mukherjee", "Arijit Ukil", "Shalini Mukhopadhyay"),
}
SOURCES = {
    "institution": "OpenAlex institution",
    "raw_affiliation": "OpenAlex affiliation text",
    "author_snowball": "Author snowball",
    "author_snowball+crossref": "Snowball Crossref follow-up",
    "dblp": "DBLP via SPARQL",
    "crossref": "Crossref affiliation search",
    "arxiv": "arXiv",
}


def rows(book, sheet, fields):
    entries = book[sheet].values
    headers = next(entries)
    cols = {field: headers.index(field) for field in fields}
    return [{field: row[index] for field, index in cols.items()}
            for row in entries if any(value is not None for value in row)]


def ranked(values, limit=None):
    return sorted(Counter(values).items(), key=lambda item: (-item[1], item[0]))[:limit]


def source_stats(papers):
    coverage = papers["Coverage"].values
    for row in coverage:
        if row[:3] == ("source", "works_returned", "new_unique"):
            break
    else:
        raise ValueError("Coverage source-pass header missing")

    additions = Counter()
    benchmark = unconfirmed = None
    for source, returned, added, note, *_ in coverage:
        if not source:
            break
        key = source.split(":", 1)[0]
        additions[key] += added
        if key == "institution":
            benchmark = returned
        if key == "author_snowball+crossref":
            unconfirmed = int(re.search(r"([\d,]+) candidates checked", note)[1].replace(",", ""))

    if benchmark is None or unconfirmed is None:
        raise ValueError("Coverage benchmark or candidate aggregate missing")
    return [(label, additions[key]) for key, label in SOURCES.items()], benchmark, unconfirmed


def group_gap(authors, directory):
    names = {row["name"] for row in directory}
    if any(set(group) - names for group in GROUPS.values()):
        raise ValueError("A comparison group member is missing from Author_Directory")

    graph = nx.Graph()
    labels = {}
    fields = ("author_a", "author_b", "author_a_id", "author_b_id")
    edges = rows(authors, "Coauthor_Network", fields)
    for edge in sorted(edges, key=lambda edge: (edge["author_a_id"], edge["author_b_id"])):
        a, b = edge["author_a_id"], edge["author_b_id"]
        graph.add_edge(a, b)
        labels.update({a: edge["author_a"], b: edge["author_b"]})

    # The export prefixes resolved directory identities with P:.
    paths = []
    for start in GROUPS["unlearning"]:
        for end in GROUPS["edge"]:
            a, b = f"P:{start}", f"P:{end}"
            if a in graph and b in graph and nx.has_path(graph, a, b):
                paths.append(nx.shortest_path(graph, a, b))
    path = min(paths, key=lambda path: (len(path), tuple(labels[node] for node in path)), default=[])
    return {
        "groups": GROUPS,
        "shared_papers": len(rows(authors, "Group_Overlap", ("title",))),
        "path": [labels[node] for node in path],
        "path_hops": len(path) - 1 if path else None,
    }


def paper_link(row):
    url = row["arxiv_url"] or row["doi"]
    if not url:
        raise ValueError(f"Missing arXiv/DOI link for {row['title']}")
    parsed = urlsplit(url)
    if parsed.scheme not in {"https", "http"} or parsed.netloc not in {
        "arxiv.org", "www.arxiv.org", "doi.org", "dx.doi.org"
    }:
        raise ValueError(f"Unexpected work link for {row['title']}")
    return url


def domain_trends(records):
    """Count rows once per domain within each explicitly defined population."""
    trends = {}
    for year in sorted({row["year"] for row in records}):
        annual = [row for row in records if row["year"] == year]
        dated = []
        incomplete = 0
        for row in annual:
            value = str(row["date"] or "")
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
                incomplete += 1
                continue
            day = date.fromisoformat(value)
            if day.year != year:
                raise ValueError("Publication year and date disagree")
            if day <= date(year, 11, 23):
                dated.append(row)
        populations = {
            "annual": annual,
            "matched": dated,
            "published": [row for row in dated if row["status"] == "Published"],
        }
        trends[str(year)] = {"incomplete_dates": incomplete}
        for key, population in populations.items():
            trends[str(year)][key] = {
                "total": len(population),
                "domains": dict(Counter(row["domain"] or "Unspecified" for row in population)),
                "llm_titles": sum(bool(re.search(r"large language model|\bLLMs?\b", row["title"], re.I))
                                  for row in population),
            }
    return trends


def percent(count, total):
    value = (Decimal(count) * 100 / Decimal(total)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
    return f"{value}%"


def direction(data):
    trends = data["trends"]
    years = sorted(trends)
    domains = [name for name, _ in data["domains"]]
    modes = [("matched", "Same dates each year"), ("annual", "All recorded works by year"),
             ("published", "Same dates, published works only")]
    sections = []
    for mode, label in modes:
        series = [trends[year][mode] for year in years]
        headers = "".join(f'<th scope="col">{year}</th>' for year in years)
        body = []
        for domain in domains:
            cells = []
            for population in series:
                count = population["domains"].get(domain, 0)
                share = percent(count, population["total"])
                cells.append(f'<td>{count}<small>{share}</small>'
                             f'<span class="mini-track" aria-hidden="true"><i style="width:{share}"></i></span></td>')
            body.append(f'<tr><th scope="row">{h(domain)}</th>{"".join(cells)}</tr>')
        totals = "".join(f'<td>{p["total"]}</td>' for p in series)
        expanded = " open" if mode == "matched" else ""
        sections.append(f'<details data-window="{mode}"{expanded}><summary>{label}</summary><table class="trends">'
                        f'<thead><tr><th scope="col">Domain</th>{headers}</tr></thead>'
                        f'<tbody><tr><th scope="row">All domains</th>{totals}</tr>{"".join(body)}</tbody></table></details>')
    matched = [trends[y]["matched"] for y in years]
    def shares(domain):
        return ", ".join(percent(p["domains"].get(domain, 0), p["total"]) for p in matched)
    recent_published = [trends[y]["published"] for y in years[-2:]]
    published_shares = [percent(p["domains"].get("NLP/LLM", 0), p["total"]) for p in recent_published]
    title_shares = ", ".join(percent(p["llm_titles"], p["total"]) for p in matched)
    direct = [p["domains"].get("Agents", 0) for p in matched]
    pub_agents = trends[years[-1]]["published"]["domains"].get("Agents", 0)
    undated = sum(trends[y]["incomplete_dates"] for y in years)
    names = ["NLP/LLM", "CV/VLM", "Edge/Systems", "Agents", "Health/Bio", "Quantum"]
    peak = max(100 * p["domains"].get(name, 0) / p["total"] for name in names for p in matched)
    ceiling = max(5, 5 * math.ceil(peak / 5))
    plots = []
    for name in names:
        values = [100 * p["domains"].get(name, 0) / p["total"] for p in matched]
        points = [(30 + i * 164 / (len(values) - 1), 68 - v * 57 / ceiling)
                  for i, v in enumerate(values)]
        vertices = " ".join(f"{x:.1f},{y:.1f}" for x, y in points)
        dots = "".join(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3"/>' for x, y in points)
        plots.append(f'<figure><figcaption>{h(name)}</figcaption>'
                     f'<svg viewBox="0 0 204 90" role="img" aria-label="{h(name)}, shares from {years[0]} '
                     f'through {years[-1]}: {shares(name)}"><path d="M30 11H194 M30 68H194"/>'
                     f'<text x="0" y="15">{ceiling}%</text><text x="0" y="71">0%</text>'
                     f'<polyline points="{vertices}"/>{dots}<text x="30" y="85">{years[0]}</text>'
                     f'<text x="194" y="85" text-anchor="end">{years[-1]}</text></svg></figure>')
    chart = '<h3>Share of clean works through 23 November each year</h3><div class="trend-plots">' + "".join(plots) + '</div>'
    return (f'<p><strong>Language models gained share throughout the window.</strong> '
            f'On the same calendar dates each year, NLP/LLM accounted for {shares("NLP/LLM")} '
            f'in chronological order. A separate title check for "large language model", "LLM" or "LLMs" '
            f'gave {title_shares} of all eligible works. Among published works alone, the NLP/LLM '
            f'share was nearly flat between {years[-2]} ({published_shares[0]}) and '
            f'{years[-1]} ({published_shares[1]}). The recent increase is more visible when preprints are included.</p>'
            f'<p><strong>Agent research rose most recently.</strong> The comparable yearly counts were '
            f'{", ".join(map(str, direct))}. Of the {direct[-1]} in the last window, {pub_agents} were marked published. '
            'The category includes multi-agent methods as well as language-model agents.</p>'
            f'<p><strong>The shift spans several fields.</strong> Edge/Systems accounted for {shares("Edge/Systems")} '
            f'across the four windows. Quantum accounted for {shares("Quantum")}. '
            'CV/VLM combines computer vision and vision-language research, so its total cannot isolate multimodal model growth.</p>'
            '<p>Each domain cell in the tables gives the paper count and its share of that column. Table bars use a common 0 to 100% scale. '
            f'The same-date views cover 1 January through 23 November and exclude {undated} clean records '
            'without a complete date. The annual view includes them and covers only the available part of 2026. '
            'The published-only view also excludes preprints and uses the remaining works as its denominator.</p>'
            + chart + "".join(sections))


def analyze(dataset):
    papers = load_workbook(dataset / "tcs_papers_2023_2026.xlsx", read_only=True, data_only=True)
    authors = load_workbook(dataset / "tcs_authors_papers.xlsx", read_only=True, data_only=True)
    try:
        all_years = Counter(row["year"] for row in rows(papers, "All", ("year",)))
        clean = rows(papers, "TCS_Research_Clean", (*CSV_FIELDS, "date"))
        clean_years = Counter(row["year"] for row in clean)
        tiers = ranked(row["venue_tier"] for row in rows(papers, "Top_Tier", ("venue_tier",)))
        sources, benchmark, unconfirmed = source_stats(papers)
        directory = rows(authors, "Author_Directory", ("name", "tcs_city", "top_3_domains", "papers"))
        gap = group_gap(authors, directory)
        collaborators = rows(authors, "External_Collaborators", ("institution", "shared_papers"))
    finally:
        papers.close()
        authors.close()

    focus = []
    for name, intro, domain, pattern in FOCUS:
        regex = re.compile(pattern, re.I) if pattern else None
        matches = (row for row in clean
                   if row["domain"] == domain or (regex and regex.search(row["title"])))
        ordered = sorted(matches, key=lambda row: (-row["year"], row["title"].casefold()))
        focus.append({"name": name, "intro": intro, "papers": [
            {key: row[key] for key in ("title", "year", "venue", "status")} | {"url": paper_link(row)}
            for row in ordered
        ]})

    with (dataset / "verification_sample.csv").open(encoding="utf-8-sig", newline="") as file:
        verdicts = Counter(row["verdict"] for row in csv.DictReader(file))
    # The author-export spot check survives as aggregates in the dataset documentation.
    readme = (dataset / "DATA_README.md").read_text(encoding="utf-8")
    author_check = re.search(r"\| Author export, (\d+) rows \| (\d+) \| (\d+) \| (\d+) \|", readme)
    if not author_check:
        raise ValueError("Author-export spot-check aggregate missing")
    checked, confirmed, contradicted, unverifiable = map(int, author_check.groups())
    if checked != confirmed + contradicted + unverifiable:
        raise ValueError("Author-export spot-check totals disagree")

    grouped = defaultdict(list)
    for row in sorted(directory, key=lambda row: (-row["papers"], row["name"].casefold())):
        grouped[row["tcs_city"] or "City unspecified"].append({
            "name": row["name"], "topics": (row["top_3_domains"] or "").split("; ") if row["top_3_domains"] else [],
            "papers": row["papers"],
        })

    return {
        "data_date": "2026-11-23",
        "dataset_version": "v1.1",
        "all": all_years.total(),
        "clean": len(clean),
        "top_tier": sum(count for _, count in tiers),
        "tiers": tiers,
        "authors_tracked": len(directory),
        "years": [{"year": year, "all": all_years[year], "clean": clean_years[year]}
                  for year in sorted(all_years | clean_years)],
        "trends": domain_trends(clean),
        "sources": sources,
        "openalex_benchmark": benchmark,
        "unconfirmed_candidates": unconfirmed,
        "domains": ranked(row["domain"] or "Unspecified" for row in clean),
        "venues": ranked((row["venue"] for row in clean if row["venue"]), 15),
        "cities": dict(sorted(grouped.items(), key=lambda item: (item[0] == "City unspecified", item[0]))),
        "focus": focus,
        "gap": gap,
        "external_institutions": sorted(collaborators, key=lambda row: (-row["shared_papers"], row["institution"]))[:10],
        "spot_checks": {
            "checked": verdicts.total() + checked,
            "confirmed": verdicts["confirmed"] + confirmed,
            "contradicted": verdicts["contradicted"] + contradicted,
            "unverifiable": verdicts["unverifiable_by_checker"] + unverifiable,
        },
    }, clean


def h(value):
    return escape(str(value))


def number(value):
    return f"{value:,}"


def bar(label, value, scale, css_class=""):
    return (f'<div class="bar-row {css_class}"><span>{h(label)}</span>'
            f'<span class="bar-track" aria-hidden="true"><span style="--size:{100 * value / scale:.4f}%"></span></span>'
            f'<strong>{number(value)}</strong></div>')


def count_table(entries, label):
    body = "".join(f'<tr><th scope="row">{h(name)}</th><td>{number(count)}</td></tr>' for name, count in entries)
    return (f'<table><thead><tr><th scope="col">{h(label)}</th><th scope="col">Works</th></tr></thead>'
            f'<tbody>{body}</tbody></table>')


def city_tables(cities):
    sections = []
    for city, researchers in cities.items():
        body = "".join(
            f'<tr><th scope="row">{h(row["name"])}</th><td>{h(", ".join(row["topics"]) or "Unspecified")}</td>'
            f'<td>{number(row["papers"])}</td></tr>' for row in researchers
        )
        sections.append(
            f'<details><summary>{h(city)} <span class="meta">· {len(researchers)} researchers</span></summary>'
            '<table><thead><tr><th scope="col">Researcher</th><th scope="col">Main topics</th>'
            f'<th scope="col">Papers</th></tr></thead><tbody>{body}</tbody></table></details>'
        )
    return "\n".join(sections)


def focus_lists(focus):
    sections = []
    for area in focus:
        entries = "".join(
            f'<li data-year="{p["year"]}" data-status="{h(p["status"])}"><a href="{h(p["url"])}">{h(p["title"])}</a><span class="meta paper-meta">'
            f'{p["year"]} · {h(p["venue"] or "Venue unspecified")} · {h(p["status"])}</span></li>'
            for p in area["papers"]
        )
        sections.append(
            f'<div class="focus-area"><h3>{h(area["name"])}</h3><p>{h(area["intro"])}</p>'
            f'<details><summary>Browse {len(area["papers"])} papers</summary><ul class="papers">{entries}</ul></details></div>'
        )
    return "\n".join(sections)


def path_diagram(path):
    if not path:
        return "<p>No connecting path appears in the coauthor network.</p>"
    height = 40 + 38 * (len(path) - 1)
    label = h(" → ".join(path))
    nodes = "".join(
        f'<circle cx="15" cy="{20 + index * 38}" r="5"/>'
        f'<text x="33" y="{25 + index * 38}">{h(name)}</text>'
        for index, name in enumerate(path)
    )
    return (f'<svg class="path" viewBox="0 0 320 {height}" role="img" aria-label="{label}">'
            f'<line x1="15" y1="20" x2="15" y2="{height - 20}"/>{nodes}</svg>')


def render(data):
    years = data["years"]
    scale = max(year["all"] for year in years)
    output = "\n".join(
        f'<div class="year"><h3>{year["year"]}</h3>'
        + bar("All", year["all"], scale, "all")
        + bar("Clean", year["clean"], scale) + "</div>" for year in years
    )
    tier_breakdown = " · ".join(f"{count} {tier}" for tier, count in data["tiers"])
    largest = [name for name, _ in data["domains"] if name != "Other"][:2]
    gap = data["gap"]
    path_hops = gap["path_hops"]
    path_length = (f"{path_hops} coauthor {'link' if path_hops == 1 else 'links'}"
                   if path_hops is not None else "no connecting path")
    gap_sentence = f"The two comparison groups share {number(gap['shared_papers'])} papers in the author export"
    if path_hops is not None:
        gap_sentence += f" and their shortest path spans {path_length}"
    summary = (
        f"The dataset contains {number(data['all'])} unique works, including {number(data['clean'])} "
        f"in the clean TCS Research subset. The tiered shortlist has {number(data['top_tier'])} works "
        f"and the author export tracks {number(data['authors_tracked'])} researchers. "
        f"{largest[0]} and {largest[1]} are the largest named domains in the clean subset. {gap_sentence}."
    )
    metadata = json.dumps({
        "@context": "https://schema.org", "@type": "ScholarlyArticle", "headline": TITLE,
        "description": DESCRIPTION, "url": URL, "inLanguage": "en",
        "author": {"@type": "Person", "name": "Prashant Badiger", "url": "https://jatasyaprashant.github.io/"},
    }, ensure_ascii=False, indent=2).replace("<", "\\u003c")
    check = data["spot_checks"]
    return Template((ROOT / "tools" / "tcs_page.html").read_text(encoding="utf-8")).substitute(
        title=h(TITLE), description=h(DESCRIPTION), url=URL, json_ld=metadata,
        data_date=h(data["data_date"]), version=h(data["dataset_version"]),
        summary=h(summary), all=number(data["all"]), clean=number(data["clean"]),
        top_tier=number(data["top_tier"]), tiers=h(tier_breakdown), authors=number(data["authors_tracked"]),
        explorer_script=(ROOT / "tools" / "tcs_explorer.js").read_text(encoding="utf-8-sig"),
        direction=direction(data), output=output, benchmark=number(data["openalex_benchmark"]),
        domains="\n".join(bar(name, count, data["domains"][0][1]) for name, count in data["domains"]),
        venues=count_table(data["venues"], "Venue"),
        cities=city_tables(data["cities"]),
        institutions=count_table(((row["institution"], row["shared_papers"])
                                  for row in data["external_institutions"]), "External institution"),
        focus=focus_lists(data["focus"]),
        unlearning=h(", ".join(gap["groups"]["unlearning"])),
        edge=h(", ".join(gap["groups"]["edge"])),
        shared=number(gap["shared_papers"]),
        path_length=h(path_length), path=path_diagram(gap["path"]),
        sources=count_table(data["sources"], "Source pass"),
        checked=number(check["checked"]), confirmed=number(check["confirmed"]),
        contradicted=number(check["contradicted"]), unverifiable=number(check["unverifiable"]),
        candidates=number(data["unconfirmed_candidates"]),
    )


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("dataset", type=Path, help="Directory containing the v1.1 dataset workbooks")
    args = parser.parse_args()
    data, clean = analyze(args.dataset)
    summary = ROOT / "tools" / "tcs_summary.json"
    summary.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    page = ROOT / "tcs-research"
    (page / "index.html").write_text(render(json.loads(summary.read_text(encoding="utf-8"))), encoding="utf-8")
    export = page / "data" / "tcs_research_clean_2023_2026.csv"
    export.parent.mkdir(parents=True, exist_ok=True)
    with export.open("w", encoding="utf-8", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=CSV_FIELDS)
        writer.writeheader()
        writer.writerows({field: row[field] for field in CSV_FIELDS} for row in clean)


if __name__ == "__main__":
    main()
