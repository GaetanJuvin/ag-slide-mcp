from ag_slide_mcp.server import server


def main():
    server.run(transport="stdio")


if __name__ == "__main__":
    main()
