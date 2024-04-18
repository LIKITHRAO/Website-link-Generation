import pandas as pd
from googlesearch import search

def fetch_links_for_queries(queries):
    links = []
    for query in queries:
        try:
            # Perform Google search for the query and extract links
            search_results = search(query, num=3, stop=3)
            # Filter out links from LinkedIn and Wikipedia
            filtered_links = [link for link in search_results if "linkedin.com" not in link and "wikipedia.org" not in link]
            links.append([query, ', '.join(filtered_links)])  # Concatenate all links into a single string
        except Exception as e:
            print(f"Error fetching search results for query: {query}")
            print(e)
            links.append([query, ''])  # Fill with an empty string if error occurs
    return links

def main(excel_file, queries_column):
    # Read queries from Excel file
    try:
        input_data_df = pd.read_excel(excel_file)
        queries = input_data_df[queries_column].tolist()
    except Exception as e:
        print("Error reading Excel file:", e)
        return

    # Fetch links for queries
    links = fetch_links_for_queries(queries)

    # Create DataFrame from links
    links_df = pd.DataFrame(links, columns=["Query", "Links"])

    # Write links to Excel file
    output_file = "C:\\Users\\likit\\Downloads\\Travel_Destination_Answers01.xlsx"
    try:
        links_df.to_excel(output_file, index=False)
        print("Links saved to:", output_file)
    except Exception as e:
        print("Error writing links to Excel file:", e)

if __name__ == "__main__":
    excel_file = "C:\\Users\\likit\\Downloads\\Travel_Destination_Answers01.xlsx"  # Path to Excel file containing queries
    queries_column = "Questions"  # Column name containing queries
    main(excel_file, queries_column)
