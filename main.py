

import openpyxl

class UserStory:

    def __init__(self, id=0, story=None, reqs=None, priority=None):
        self.id =id
        self.description = story
        self.reqs = reqs
        self.priority = priority

    def __str__(self):

        mystr = '''User Story %s
        %s
        Requirements: %s
        Priority: %s
        '''
        return mystr % (str(self.id),self.description,self.reqs,self.priority)

d = openpyxl.load_workbook("test.xlsx")
s1 = d.get_sheet_by_name('User Stories')

in_data = False

for r in s1.rows:

    if in_data:
        my_story = UserStory()
        my_story.id = r[0].value
        my_story.description = r[1].value
        my_story.reqs = r[2].value
        my_story.priority = r[3].value

        print my_story

    if r[0].value is not None and r[1].value is not None and r[2].value is not None and r[3].value is not None:
        in_data = True

