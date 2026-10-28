from dotenv import load_dotenv
load_dotenv()

from src.sharepoint_uploader import upload_to_sharepoint
upload_to_sharepoint("output/mynewsdesk_metrics.xlsx")