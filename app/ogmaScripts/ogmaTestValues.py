import datetime

def get_current_datetime_str() -> str:
    # for testing, can be deleted.
    # Get the current datetime
    now: datetime.datetime = datetime.datetime.now()

    # Format the datetime string as YYYYMMDD-HH.MM[AM|PM]
    formatted_datetime: str = now.strftime("%Y%m%d-%I.%M%p")

    return formatted_datetime


def modify_word_properties(file_paths: list[str] | str, properties: dict[str, str] | None = None) -> tuple[list[str],dict[str,str]]:
    # Define the properties and their default values

    # If passed a single file, make it a list with a single index
    if isinstance(file_paths, str):
        file_paths = [file_paths]

    time: str = get_current_datetime_str()

    # If no properties are passed, default ones
    if properties == None:
        properties = {
            "BOK ID": f"BOK ID {time}",
            "Document Name": f"DOC NAME {time}",
            "Company Name": f"CO NAME {time}",
            "Division": f"DIV {time}",
            "Author": f"AUTH {time}",
            "Company Address": f"ADDR {time}",
            "Project Name": f"PRJ NAME {time}",
            "Project Number": f"PRJ ID {time}",
            "End Customer": f"END CUST {time}",
            "Site Name": f"SITE NAME {time}",
            "File Name": f"FILE NAME {time}",
        }

    return (file_paths, properties)