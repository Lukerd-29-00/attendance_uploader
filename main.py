import re
import openpyxl
import typing
import asyncio
import openpyxl.worksheet._read_only
import aiohttp
import json

BASE = "https://udel.instructure.com/api/v1" #The base URL for canvas.

COURSE = 1671413 #You can get this from the URL bar anywhere in the 220 canvas page.
ASSIGNMENT = 10570458 #you can get the assignment ID from the URL bar in speedgrader.

ATTENDANCE_PATH="input/032.xlsx" #The path to the workbook containing attendance information (see above). 
ATTENDANCE_SHEET = "test" #The name of the worksheet inside the ATTENDANCE_PATH workbook (see above).

IDS_PATH = "input/ids.xlsx" #The path to the student canvas IDs (see above).
IDS_SHEET = "Sheet1" #The name of the sheet containing the canvas IDs in the IDS_PATH workbook (see above).

AUTH_PATH = "input/attendance_upload_token.txt" #The path to your Canvas OAuth2 token (see above).

class Name(object):
    """Represents a student's name.
    
    Attributes:
        first: The first name of the student.
        last: The student's last name.
        middle: The student's middle name. Optional.
    """
    first: str
    last: str
    middle: typing.Union[None,str]
    def __init__(self,first: str,last: str,middle: typing.Union[str,None]=None):
        """Initializes a name.

        Args:
            first:
                The student's first name.
            last:
                The student's last name.
            middle:
                The student's middle name. Optional.
        """
        self.first = first.strip().lower()
        self.last = last.strip().lower()
        if isinstance(middle,str):
            self.middle = middle.strip().lower()
        else:
            self.middle = None
    def __eq__(self,other: 'Name')->bool:
        """Determines whether two students have the same name.

        Args:
            other:
                The name being compared to this one.
        
        Returns:
            bool: Whether the first, last, and middle names are the same for both students.
        """
        return self.first == other.first and self.last == other.last and self.middle == other.middle
    def __hash__(self)->int:
        """Produces a hash of this name.

        Returns:
            The hash of the first, middle (where applicable) and last names concatenated together.
        """
        if self.middle:
            return hash(self.first + self.middle + self.last)
        else:
            return hash(self.first + self.last)
    def __repr__(self)->str:
        """Creates a string representing the student's name.
        
        Returns:
            A string representing the name in middle last,first format.
        """
        if self.middle:
            return f"{self.middle.capitalize()} {self.last.capitalize()},{self.first.capitalize()}"
        else:
            return f"{self.last.capitalize()},{self.first.capitalize()}"

def create_name(raw: str)->Name:
    """Creates a Name object from a middle last,first formatted string.

    Args:
        raw:
            The raw string being converted to a Name object. It should be in the format middle last,first where the middle name is optional.
    
    Returns:
        A Name object with the appropriate first, middle, and last names from the raw string. 
    """
    match = re.search(r"([a-zA-Z-']+)\s+([a-zA-Z-']+)\s*,\s*([a-zA-Z-']+)",raw)
    if match:
        return Name(match.group(3),match.group(2),match.group(1))
    else:
        match = re.search(r"([a-zA-Z-']+)\s*,\s*([a-zA-Z-']+)",raw)
        return Name(match.group(2),match.group(1))

def get_ids(filename: str)->typing.Dict[Name,int]:
    """Extracts the canvas ID's from a worksheet. Ignores the first row

    Args:
        filename:
            The file to extract the ids from. Should have a heading in the first row, with the names in the first column and the ids in the second.
    
    Returns:
        A dictionary mapping the Name objects extracted from the first column to the ids in the second.
    """
    ids = openpyxl.load_workbook(filename=filename,read_only=True)
    idsmap: typing.Dict[Name,int] = {}
    header = True
    for row in ids[IDS_SHEET].rows:
        if header:
            header = False
            continue
        name=create_name(row[0].value)
        idsmap[name] = row[1].value
    ids.close()
    return idsmap


async def update_attendance(sheet: openpyxl.worksheet._read_only.ReadOnlyWorksheet,ids: typing.Dict[Name,int])->None:
    """Asynchronously sets the grades of the students in the provided worksheet via the Canvas REST API.

    Args:
        sheet:
            A read-only worksheet that contains any info the first column (purely for human use, such as email), the names of the students in the second, and whether or not they were present in the third. A student will be marked present if their column says present or yes, and absent if it says absent or no. The first row will be ignored.
        ids:
            A dictionary mapping the names of the students in sheet to their canvas IDs.
    """
    url = f"{BASE}/courses/{COURSE}/assignments/{ASSIGNMENT}/submissions"
    header = True
    token = ""
    with open(AUTH_PATH,"r") as f:
        token = f.read().strip()
    async with aiohttp.ClientSession() as session:
        for row in sheet:
            if header:
                header = False
                continue
            student_id = ids[create_name(row[1].value)]
            present: str = row[2].value.strip().lower()
            grade = 0
            if present == "present" or present == "yes":
                grade = 5
            elif present == "absent" or present == "no":
                grade = 0
            else:
                raise ValueError(f"Inappropriate entry {present} in entry for {create_name(row[0].value)}")
            body={
                "submission": {
                    "assignment_id":ASSIGNMENT,
                    "posted_grade":grade,
                    "user_id":student_id,
                    "submission_type": None
                }
            }
            headers = {
                "Content-Type": "application/json",
                "Accept": "application/json",
                "Authorization": f"Bearer {token}"
            }
            async with session.put(f"{url}/{student_id}",data=json.dumps(body),headers=headers) as res:
                if not res.ok:
                    raise ValueError(await res.text())

async def main():
    """Gets the IDs from the appropriate sheet, reads the attendance sheet, and marks the students as present or absent appropriately."""
    ids = get_ids(IDS_PATH)
    wb = openpyxl.load_workbook(filename=ATTENDANCE_PATH,read_only=True)
    testSheet: openpyxl.worksheet._read_only.ReadOnlyWorksheet = wb[ATTENDANCE_SHEET]
    await update_attendance(testSheet,ids)
    wb.close()

if __name__ == "__main__":
    #uses an outdated syntax due to a bug with asyncio.run on Windows. Ignore the DeprecationWarning.
    asyncio.get_event_loop().run_until_complete(main())