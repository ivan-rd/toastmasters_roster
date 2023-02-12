# toastmasters_roster

The purpose of this project is to produce a toastmasters club roster with
automated role assignment. There are two files in the source directory.
roster_processing.py is used to produce a new roster.
merge_rosters.py is used to append the new roster to the current roster.
No arguments are passed to these Python scripts when they are called.
However, they read a number of spreadsheets, which must be placed
in the source directory.

The minimum version of Python is 3.10 since switch-case statements are used.
The required modules pandas, numpy, yaml, with their dependencies.

1. Set input spreadsheets

meeting_availability.xlsx (row index: members, column index: dates)

New meeting dates are basically set in the columns in of this spreadsheet.

role_availability.xlsx (row index: members, column index: roles )

The row index in this spreadsheet must match the row index of meeting_availability.xslx
There are no dates - the role set is constant for every member during the
period spanned by the new roster. Besides the meeting roles, the last column, labeled 'Multirole level',
assigns one of the three levels to each member. This is relevant when the number
of meeting slots exceeds the number of available members, leading to multiple
roles for some members. There are three possible levels: beginner, intermediate
and advanced.

role_schedule.xlsx (row index: meeting slots, column index: dates)

The column index must match the column index of meeting_availability.xlsx
This spreadsheet allows the user to turn off some roles as needed
for every meeting.

double_roles.xlsx (row index: roles, column index: roles)

When setting multiple roles, some pairs of roles are allowed and others
are not, as defined in this spreadsheet. The settings are different
depending on the 'Multirole level', which is set in role_availability.xlsx
for every member. Hence, the file contains three separate sheets labelled
'beginner', 'intermediate' and 'advanced. Row and column indices are the same -
every role is listed in both. Consequently, there is effectively a double
entry for each pair. The table should be symmetrical about the diagonal.

role_spacing.xlsx (row index: roles, column index: role spacing)

There are only two columns: the first column lists roles and the second
column has numbers showing a minimum time spacing for each role i.e.
a minimum number of meetings before a member can be assigned the same
role again.

club_roster.xlsx (row index: members, columns index: dates)

No action is needed here. roster_processing.py reads the roster history
to work out time intervals between roles and role frequency. However,
it can also start from an empty roster.

setting.yaml

A YAML file that contains a few parameters. Normally, this would only
be used by developers and advanced users.

The input_files_examples directory contains examples of input spreadsheets
for three clubs with a varying number of members. The medium club scenario
,where the number of members is around 20, is probably easiest for rostering.
In the small club scenario (around 10 members), the challenge is to manage
the burden of multitasking with multiple roles. Conversely, in a large club
(around 30 members), the challenge is to avoid having a streak of meetings
where some members are not assigned any roles at all.

It is also possible to start with an empty club roster, which is given
in input_files_examples/empty_roster.

All member names in the input spreadsheets should be spelled correctly as
there is not any kind of string checking. A common pitfall is to leave
blank spaces at the end. This would result in a different name to the
intended one.

2. Run roster_processing.py

A new roster is printed out and saved in new_roster.xlsx. Any meeting slots
that could not be assigned due to constraints in input spreadsheets are also
printed out to provide feedback to the user. Besides a new roster, there is also another
sheet new_roster.xlsx which shows what roles each member was assigned. If a new roster is
satisfactory go to step 3. Otherwise go back to step 1 and modify input spreadsheets
to allow more freedom in assigning, or reduce the number of meeting slots.

3. Run merge_rosters.py

This script appends new_roster.xlsx to club_roster.xlsx. The list of members
in the two spreadsheets does not have to be the same. Meeting entries for
members who do not appear in the new_roster.xlsx are left blank in the club_roster.xlsx.
