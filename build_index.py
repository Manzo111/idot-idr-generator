"""Build or refresh the persistent IDOT contract lookup index."""

import argparse

import app


def parse_args():
    parser = argparse.ArgumentParser(
        description=(
            "Download visible IDOT letting pages once and store contract URLs in "
            "idot_contract_index.sqlite for fast Streamlit lookups."
        )
    )
    parser.add_argument(
        "--recent",
        type=int,
        default=0,
        metavar="N",
        help="Only index the N newest letting pages instead of the full visible archive.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Refresh pages even when the saved archive signature is still fresh.",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    session = app.make_session()
    letting_links = app.get_all_archive_letting_links_newest_first(session)

    if args.recent > 0:
        letting_links = letting_links[: args.recent]
        result = app.build_sqlite_contract_index_for_lettings(
            letting_links,
            max_pages_per_letting=1,
            meta_key=f"cli_recent_{args.recent}",
            meta_ttl_seconds=app.RECENT_INDEX_TTL_SECONDS,
            force=args.force,
            source_suffix="cli-recent-index",
        )
    else:
        result = app.ensure_full_sqlite_contract_index(
            letting_links,
            force=args.force,
        )

    stats = app.sqlite_index_stats()

    print(f"Visible lettings discovered: {len(letting_links)}")
    print(f"Lettings checked now: {result.get('checked_lettings', 0)}")
    print(f"Contract rows saved now: {result.get('saved_contracts', 0)}")
    print(f"Failed letting requests now: {result.get('failed_lettings', 0)}")
    print(
        "Persistent index totals: "
        f"{stats['contracts']} contracts, "
        f"{stats['indexed_lettings']}/{stats['known_lettings']} lettings indexed."
    )
    print(f"Database: {app.CONTRACT_INDEX_DB_PATH}")


if __name__ == "__main__":
    main()
