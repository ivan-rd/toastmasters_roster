import pandas as pd

roster_current_dict = pd.read_excel("club_roster.xlsx",sheet_name=["Current Roster","Members"])
roster_appended_dict = pd.read_excel("new_roster.xlsx",sheet_name=["Current Roster","Members"])
roster_current = roster_current_dict["Current Roster"]
members_current = roster_current_dict["Members"]
roster_appended = roster_appended_dict["Current Roster"]
members_appended = roster_appended_dict["Members"]
roster_current = roster_current.set_index("Meeting Date:")
members_current = members_current.set_index("Members")
roster_appended = roster_appended.set_index("Meeting Date:")
members_appended = members_appended.set_index("Members")
roster_current.columns = list(map(lambda x:x.date(),roster_current.columns.values))
members_current.columns = list(map(lambda x:x.date(),members_current.columns.values))
roster_appended.columns = list(map(lambda x:x.date(),roster_appended.columns.values))
members_appended.columns = list(map(lambda x:x.date(),members_appended.columns.values))

roster_new = pd.concat([roster_current, roster_appended], axis=1)
members_new = pd.concat([members_current, members_appended], axis=1)  

writer = pd.ExcelWriter('club_roster.xlsx', engine='xlsxwriter')
roster_new.to_excel(writer, sheet_name='Current Roster')
members_new.to_excel(writer, sheet_name='Members')
writer.save()
