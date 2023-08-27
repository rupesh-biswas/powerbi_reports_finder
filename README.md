# Power BI API Data Retrieval Script

This Python script interacts with the Power BI REST API to retrieve information about workspaces, reports, and data sources from your Power BI account. The retrieved data is then saved into an Excel file for further analysis.

## Prerequisites

Before running the script, ensure you have the following:

- Python 3.x installed on your system.
- Required libraries: `requests`, `pandas`, `webbrowser`.

## Usage

1. Clone this repository to your local machine:

   ```bash
   git clone https://github.com/rupesh-biswas/powerbi_reports_finder.git
   ```

2. Install the required libraries:

   ```bash
   pip install requests pandas
   ```

3. Run the script:

   ```bash
   python getPowerBiDatasources.py
   ```

4. Follow the prompts to provide necessary inputs:
   - Choose admin mode (Y/N) to specify API access level. Only Azure admin can generate admin bearer token.
   - Copy the bearer token from the provided URL.
   
5. The script will retrieve information about workspaces, reports, and data sources based on your inputs.

6. Once the script finishes, you'll find an Excel file named "powerBI_datasets.xlsx" in the same location as the script.

## Configuration

- Replace the URLs (`toekn_url`, `workspaces_url`, `reports_url`, `datasource_url`) with valid URLs from the Power BI API documentation in case of any url errors.
- Modify the script to suit your specific requirements.

## License

This project is licensed under the [MIT License](LICENSE).

## Acknowledgements

- This script was inspired by the need to automate Power BI API data retrieval.
- Thanks to the Power BI API documentation for providing the necessary information to interact with the API.

---

Feel free to contribute to this project by opening issues or pull requests!

For questions or support, please contact [rupeshbiswas1997@gmail.com](mailto:rupeshbiswas1997@gmail.com).