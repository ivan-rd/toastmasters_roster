import pandas as pd
import numpy as np
import yaml
from yaml.loader import SafeLoader

# Returns a role type given a roster slot

def match_roster_role(roster_role): 
    match roster_role:
        case 'Set Up/Pack Up #1':
            return 'Set Up/Pack Up'
        case 'Set Up/Pack Up #2':
            return 'Set Up/Pack Up'
        case 'Meet & Greet #1':
            return 'Meet & Greet'
        case 'Meet & Greet #2':
            return 'Meet & Greet'
        case 'Toastmaster':
            return 'Toastmaster'
        case 'General Evaluator':
            return 'General Evaluator'
        case 'Inspiration':
            return 'Inspiration'
        case 'Grammarian':
            return 'Grammarian'
        case 'Um & Ah Counter':
            return 'Um & Ah Counter'
        case 'Speaker #1':
            return 'Speaker'
        case 'Speaker #2':
            return 'Speaker'
        case 'Speaker #3':
            return 'Speaker'
        case 'Speaker #4':
            return 'Speaker'
        case 'Speaker #5':
            return 'Speaker'
        case 'Evaluator #1':
            return 'Evaluator'
        case 'Evaluator #2':
            return 'Evaluator'
        case 'Evaluator #3':
            return 'Evaluator'
        case 'Evaluator #4':
            return 'Evaluator'
        case 'Evaluator #5':
            return 'Evaluator'
        case 'Tabletopic Master': 
            return 'Table Topics Master'
        case 'TTopic Eval #1': 
            return 'Table Topics Evaluator'
        case 'TTopic Eval #2': 
            return 'Table Topics Evaluator'
        case 'Business': 
            return 'Business'
        case 'Timekeeper': 
            return 'Timekeeper'
        case 'Timekeeper Report': 
            return 'Timekeeper'

# Finds the member with the greatest time distance for a given role
# and indicates if the distance is greater than the minimum allowed
# Multiple members are returned when the maximum distance is shared.

def get_maximum_distance(distance_table,member_group,column,lower_limit): 
    distances = np.zeros(member_group.size)
    for i in range(member_group.size): 
        distances[i] = distance_table.at[member_group[i],column]
    index_positions = distances == distances.max()
    max_distance_names = member_group[index_positions]
    if distances.max() > lower_limit[column]:
        above_limit = True
    else: 
        above_limit = False
    return [max_distance_names,above_limit]

# Returns a members with a minimum role frequency. In this context 
# frequency is an integer number indicating how many times a role
# was done within a window of time.

def get_minimum_frequency(frequency_table,column,max_distance_names):
    freq_array = np.zeros(len(max_distance_names))
    for i in range(max_distance_names.size): 
        freq_array[i] = frequency_table.at[max_distance_names[i],column]
    index_positions = np.where(freq_array == freq_array.min())
    min_frequency_names = max_distance_names[index_positions]
    return min_frequency_names

# Processes the meeting availability spreadsheet to find available
# members for a particular meeting

def get_availability(column,availability_spreadsheet):
    column_list = availability_spreadsheet[column].to_numpy()
    available_members = availability_spreadsheet.index[column_list == 'Y']
    return available_members

# This is relevant when assigning multiple roles in the third pass. 
# Role affinity is defined in the weight matrix. For example, Um & Ah count 
# is more likely to be assigned to Grammarian than General Evaluator
# due to a lower weight. 

def get_weights(weights_table,assign_table,available_members,role):
    weights = weights_table.loc[role,:].values
    available_members_weights = np.zeros(len(available_members))
    members_dict = {}
    for i in range(available_members.size): 
        available_members_weights[i] = np.sum(np.multiply(assign_table.loc[available_members[i],:].values,weights))
        members_dict[available_members[i]] = available_members_weights[i]
    members_dict_sorted = sorted(members_dict.items(), key=lambda x: x[1])
    members_sorted = []
    for k in members_dict_sorted:
        members_sorted.append(k[0])
    return members_sorted

# Loads paramters from a yaml file that is likely to be used only 
# by developers and advanced users 

def get_parameters(): 
    with open('settings.yaml') as f:
        settings = yaml.load(f, Loader=SafeLoader)
    role_spacing = pd.read_excel("role_spacing.xlsx")
    role_spacing = role_spacing.set_index("Roles")
    distance_thresholds = role_spacing.to_dict()["Spacing"]

    DEBUG = True
    LOG = True
    if LOG: 
        f = open("assign_detail.txt","w")
    else:
        f = None

    #################################   spe, tme, gev, ttm, eva, tte, ins, grm, uhm, tmr, mgr, spu, num
    data_weigths_table =     np.array([[999, 80 , 80 , 75 , 90 , 70 , 80 , 70 , 70 , 80 , 20 , 100, 100],
                                       [80 , 999, 80 , 70 , 70 , 70 , 50 , 60 , 60 , 80 , 40 , 100, 90 ],
                                       [80 , 80 , 999, 90 , 90 , 90 , 70 , 90 , 90 , 90 , 30 , 100, 90 ],
                                       [75 , 70 , 90 , 999, 50 , 90 , 30 , 40 , 40 , 70 , 20 , 100, 80 ],
                                       [90 , 70 , 90 , 50 , 999, 40 , 20 , 30 , 30 , 50 , 10 , 100, 70 ],
                                       [70 , 70 , 90 , 90 , 40 , 999, 20 , 30 , 30 , 70 , 10 , 100, 70 ],
                                       [80 , 50 , 70 , 30 , 20 , 20 , 999, 20 , 20 , 30 , 10 , 100, 70 ],
                                       [70 , 60 , 90 , 40 , 30 , 30 , 20 , 999, 10 , 50 , 10 , 100, 70 ],
                                       [70 , 60 , 90 , 40 , 30 , 30 , 20 , 10 , 999, 50 , 10 , 100, 70 ],
                                       [80 , 80 , 90 , 70 , 50 , 70 , 30 , 50 , 50 , 999, 10 , 100, 80 ],
                                       [20 , 10 , 10 , 5  , 5  , 5  , 5  , 5  , 5  , 5  , 999, 100, 0  ],
                                       [100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 999, 100]])                              

    return [settings["maximum_length"], distance_thresholds, settings["distance_threshold_init"], 
            data_weigths_table, settings["LOG"], settings["DEBUG"], f, settings["second_pass_iterations"]]

def process_inputs(maximum_length): 
    roster = pd.read_excel("club_roster.xlsx")
    role_availability = pd.read_excel("role_availability.xlsx")
    meeting_availability = pd.read_excel("meeting_availability.xlsx")
    role_schedule = pd.read_excel("role_schedule.xlsx")

    roster_trimmed = roster.set_index('Meeting Date:')  
    if roster_trimmed.shape[1] >= maximum_length:
        input_roster_length = maximum_length
    else:
        input_roster_length = roster_trimmed.shape[1]
    roster_trimmed = roster_trimmed.iloc[:,roster_trimmed.shape[1]-input_roster_length:roster_trimmed.shape[1]]
    roster_trimmed.columns = list(range(input_roster_length,0,-1))

    role_availability = role_availability.set_index('Members')
    member_level = role_availability["Multirole level"]
    del role_availability["Multirole level"]
    meeting_availability = meeting_availability.set_index('Members')
    meeting_availability.columns = list(map(lambda x:x.date(),meeting_availability.columns.values))

    role_schedule = role_schedule.set_index("Meeting Date:")
    role_schedule.columns = list(map(lambda x:x.date(),role_schedule.columns.values))
    role_schedule_dict = {key: [] for key in role_schedule.columns.values}   
    for i in role_schedule.columns.values:
        for k in role_schedule.index.values:
            if role_schedule.at[k,i] == 'Y':
                role_schedule_dict[i].append(k)

    double_roles = pd.read_excel("double_roles.xlsx", sheet_name=['beginner','intermediate','advanced'])
    for i in double_roles.keys(): 
        double_roles[i] = double_roles[i].set_index('Roles')

    return [roster_trimmed, input_roster_length,role_availability, meeting_availability, member_level, double_roles, role_schedule_dict]

# The main purpose of create_tables is to go through the current roster and build a data frame which indicates how long ago 
# various roles were done by members (distance table) and how many times (frequency table) 

def create_tables(meeting_availability, role_availability, distance_threshold_init, input_roster_length, roster_trimmed, data_weights_table):

    init_data_dt = np.full((role_availability.index.size,role_availability.columns.size),input_roster_length + distance_threshold_init)
    init_data_ft = np.zeros([role_availability.index.size,role_availability.columns.size])
    distance_table = pd.DataFrame(data = init_data_dt,index = role_availability.index, columns = role_availability.columns)
    frequency_table = pd.DataFrame(data = init_data_ft,index = role_availability.index, columns = role_availability.columns)
    roles_assigned = pd.DataFrame(index = meeting_availability.index, columns = meeting_availability.columns)
    roles_assigned.iloc[:,:] = ""

    for i in roster_trimmed.columns:
        s1 = roster_trimmed.at['Speaker #1',i]
        s2 = roster_trimmed.at['Speaker #2',i]
        s3 = roster_trimmed.at['Speaker #3',i]
        s4 = roster_trimmed.at['Speaker #4',i]
        s5 = roster_trimmed.at['Speaker #5',i]
        S = [s1,s2,s3,s4,s5]
        for k in S:
            if k in distance_table.index:
                distance_table.at[k,'Speaker'] = i
                frequency_table.at[k,'Speaker'] = frequency_table.at[k,'Speaker'] + 1
        s6 = roster_trimmed.at["Toastmaster",i]
        if s6 in distance_table.index:
            distance_table.at[s6,'Toastmaster'] = i
            frequency_table.at[s6,'Toastmaster'] = frequency_table.at[s6,'Toastmaster'] + 1
        s7 = roster_trimmed.at['General Evaluator',i]
        if s7 in distance_table.index:
            distance_table.at[s7,'General Evaluator'] = i
            frequency_table.at[s7,'General Evaluator'] = frequency_table.at[s7,'General Evaluator'] + 1
        s8 = roster_trimmed.at['Tabletopic Master',i]
        if s8 in distance_table.index:
            distance_table.at[s8,'Table Topics Master'] = i
            frequency_table.at[s8,'Table Topics Master'] = frequency_table.at[s8,'Table Topics Master'] + 1
        s9 = roster_trimmed.at['Evaluator #1',i]
        s10 = roster_trimmed.at['Evaluator #2',i]
        s11 = roster_trimmed.at['Evaluator #3',i]
        s12 = roster_trimmed.at['Evaluator #4',i]
        s13 = roster_trimmed.at['Evaluator #5',i]
        S = [s9,s10,s11,s12,s13]
        for k in S:
            if k in distance_table.index:
                distance_table.at[k,'Evaluator'] = i
                frequency_table.at[k,'Evaluator'] = frequency_table.at[k,'Evaluator'] + 1
        s14 = roster_trimmed.at['TTopic Eval #1',i]
        s15 = roster_trimmed.at['TTopic Eval #2',i]
        S = [s14,s15]
        for k in S:
            if k in distance_table.index:
                distance_table.at[k,'Table Topics Evaluator'] = i
                frequency_table.at[k,'Table Topics Evaluator'] = frequency_table.at[k,'Table Topics Evaluator'] + 1
        s16 = roster_trimmed.at['Inspiration',i]
        if s16 in distance_table.index:
            distance_table.at[s16,'Inspiration'] = i
            frequency_table.at[s16,'Inspiration'] = frequency_table.at[s16,'Inspiration'] + 1
        s17 = roster_trimmed.at['Grammarian',i]
        if s17 in distance_table.index:
            distance_table.at[s17,'Grammarian'] = i
            frequency_table.at[s17,'Grammarian'] = frequency_table.at[s17,'Grammarian'] + 1
        s18 = roster_trimmed.at['Um & Ah Counter',i]
        if s18 in distance_table.index:
            distance_table.at[s18,'Um & Ah Counter'] = i
            frequency_table.at[s18,'Um & Ah Counter'] = frequency_table.at[s18,'Um & Ah Counter'] + 1
        s19 = roster_trimmed.at['Timekeeper Report',i]
        s20 = roster_trimmed.at['Timekeeper',i]
        S = [s19,s20]
        for k in S:
            if k in distance_table.index:
                distance_table.at[k,'Timekeeper'] = i
                frequency_table.at[k,'Timekeeper'] = frequency_table.at[k,'Timekeeper'] + 1
        s21 = roster_trimmed.at['Meet & Greet #1',i]
        s22 = roster_trimmed.at['Meet & Greet #2',i]
        S = [s21,s22]
        for k in S:
            if k in distance_table.index:
                distance_table.at[k,'Meet & Greet'] = i
                frequency_table.at[k,'Meet & Greet'] = frequency_table.at[k,'Meet & Greet'] + 1
        s23 = roster_trimmed.at['Set Up/Pack Up #1',i]
        s24 = roster_trimmed.at['Set Up/Pack Up #2',i]
        S = [s23,s24]
        for k in S:
            if k in distance_table.index:
                distance_table.at[k,'Set Up/Pack Up'] = i
                frequency_table.at[k,'Set Up/Pack Up'] = frequency_table.at[k,'Meet & Greet'] + 1

    index_assign_table = np.array(role_availability.index.values)
    columns_assign_table = np.array(role_availability.columns.values)
    columns_assign_table = np.append(columns_assign_table,"Number of roles")
    init_data_assign_table = np.zeros([len(index_assign_table),len(columns_assign_table)])
    assign_table = pd.DataFrame(data = init_data_assign_table,index = index_assign_table, columns = columns_assign_table)
                                  
    index_weight_table = role_availability.columns.values
    columns_weight_table = columns_assign_table
    weights_table = pd.DataFrame(data = data_weights_table,index = index_weight_table, columns = columns_weight_table)

    role_pools = {}
    for i in role_availability.columns.values: 
        role_pools[i] = get_availability(i,role_availability)

    unassigned_meetings_keys = role_availability.index.values
    unassigned_meetings = dict.fromkeys(unassigned_meetings_keys, 0)

    return [distance_table,frequency_table,assign_table,weights_table,role_pools, roles_assigned, unassigned_meetings]

def get_assigned_roles(assign_table,member):
    asssigned_roles = assign_table.columns.values[np.where(assign_table.loc[member,:])]
    return asssigned_roles

def update_unassigned_meetings(unassigned_meetings,members_not_assigned):
    for k in unassigned_meetings.keys():
        if k in members_not_assigned: 
            unassigned_meetings[k] = unassigned_meetings[k] + 1
        else: 
            unassigned_meetings[k] = 0

# Splits members according to the number of successive meetings without role assignment. All members who had a role in the last
# meeting are in one group, those who didn't are in a separate group, those who weren't assigned in for last two meetings are 
# in a separate group etc. 

def split_unassigned_meetings(unassigned_meetings): 
    unassigned_meetings_groups = {}
    for i,v in unassigned_meetings.items():
        unassigned_meetings_groups[v] = [i] if v not in unassigned_meetings_groups.keys() else unassigned_meetings_groups[v] + [i]
    return unassigned_meetings_groups

def main():

    [maximum_length, distance_thresholds, distance_threshold_init,
     data_weights_table, LOG, DEBUG, f, second_pass_iterations] = get_parameters()
    [roster_trimmed,input_roster_length, role_availability, meeting_availability, member_level, double_roles, role_schedule_dict] = process_inputs(maximum_length)
    [distance_table,frequency_table,assign_table,weights_table,role_pools, 
     roles_assigned, unassigned_meetings] = create_tables(meeting_availability,role_availability, distance_threshold_init, 
                                                          input_roster_length, roster_trimmed, data_weights_table)

    new_roster_init_data = np.empty((roster_trimmed.index.size,meeting_availability.columns.size),dtype=str)
    new_roster = pd.DataFrame(data = new_roster_init_data,index = roster_trimmed.index, columns = meeting_availability.columns)   

    for meeting_date in new_roster.columns:

        if LOG:
            print("",file=f)

        available_members_meeting = get_availability(meeting_date,meeting_availability)
        members_not_assigned = np.copy(available_members_meeting)
        unassigned_meetings_groups = split_unassigned_meetings(unassigned_meetings)

        # For every role in the currently processed meeting determine what members are available 

        role_pools_meeting = {}
        for n in role_availability.columns.values: 
            role_pools_meeting[n] = np.intersect1d(available_members_meeting,role_pools[n]) 

        role_pools_meeting_sp = role_pools_meeting.copy()  # This copy is made specifically for the second pass

        meeting_roster_assignment = role_schedule_dict[meeting_date].copy()
        assign_table.iloc[:,:] = 0
        slots_not_assigned = meeting_roster_assignment.copy()

        # FIRST PASS: 
        # Try to assign every attending member one role. The main factor when deciding who gets a role is the last time the role 
        # was done. The second factor is how frequently it was done. If ,after these two filters, mulitple members end up equal 
        # then one is selected randomly. Among members with the maximum distance score, those who haven't had a role in one or 
        # more successive meetings are given priority. 

        for slot in meeting_roster_assignment:
            role = match_roster_role(slot)  
            available_role_nas = np.intersect1d(members_not_assigned,role_pools[role])
            if len(available_role_nas) > 0:
                [max_distance_names,above_threshold] = get_maximum_distance(distance_table,available_role_nas,role,distance_thresholds) 
                if above_threshold:
                    for k in sorted(unassigned_meetings_groups.keys(),reverse=True):
                        max_distance_names_p = np.intersect1d(max_distance_names,np.array(unassigned_meetings_groups[k]))  
                        if len(max_distance_names_p) > 0:                            
                            if (len(max_distance_names_p) == 1):
                                roster_entry = max_distance_names_p[0]
                            else:
                                min_frequency_names = get_minimum_frequency(frequency_table,role,max_distance_names_p) 
                                if (len(min_frequency_names)) == 1: 
                                    roster_entry = min_frequency_names[0]
                                else: 
                                    if DEBUG: 
                                        np.random.seed(0)
                                    roster_entry = np.random.choice(min_frequency_names)
                            new_roster.at[slot,meeting_date] = roster_entry 
                            if LOG:
                                print(meeting_date," ",slot,":",roster_entry," ","First pass"," ","Distance: ",distance_table.at[roster_entry,role], file=f)
                            assign_table.at[roster_entry,role] = 1
                            assign_table.at[roster_entry,'Number of roles'] = assign_table.at[roster_entry,'Number of roles'] + 1
                            if (role == "Speaker"):
                                for k in role_pools_meeting_sp.keys(): 
                                    role_pools_meeting_sp[k] = np.delete(role_pools_meeting_sp[k],np.where(role_pools_meeting_sp[k] == roster_entry))
                            else:
                                role_pools_meeting_sp[role] = np.delete(role_pools_meeting_sp[role],np.where(role_pools_meeting_sp[role] == roster_entry))
                            members_not_assigned = np.delete(members_not_assigned,np.where(members_not_assigned == roster_entry))
                            slots_not_assigned.remove(slot)   
                            break
            if members_not_assigned.size == 0: 
                break

        if LOG:
            print(meeting_date,"  Members not assigned after first stage: ",members_not_assigned, file=f)  
            print(meeting_date,"  Slots not assigned after first stage: ",slots_not_assigned, file=f) 

        # SECOND PASS: 
        # If there are still empty slots and unassigned members, try harder to assign a role to the remaining unassigned members. 
        # This situation can especially arise with newer members whose set of roles is limited in the role availability spreadsheet. 
        # With fewer roles, their maximum time distances will be lower and their allowed roles could go to more experienced members with 
        # higher time distances. In this pass, already assigned members are pushed to other roles to free up slots that unasssigned members
        # can potentially take. 

        if (len(members_not_assigned) > 0 and len(slots_not_assigned) > 0):
            for k in range(0,second_pass_iterations):  
                if len(slots_not_assigned) == 0 or len(members_not_assigned) == 0: 
                    break
                role = match_roster_role(slots_not_assigned[0])  
                if len(role_pools_meeting_sp[role]) > 0:
                    [max_distance_names,above_threshold] = get_maximum_distance(distance_table,role_pools_meeting_sp[role],role,distance_thresholds)
                    if above_threshold:
                        if (len(max_distance_names) == 1):
                            roster_entry = max_distance_names[0]
                        else:
                            min_frequency_names = get_minimum_frequency(frequency_table,role,max_distance_names) 
                            if (len(min_frequency_names)) == 1: 
                                roster_entry = min_frequency_names[0]
                            else: 
                                if DEBUG: 
                                    np.random.seed(0)
                                roster_entry = np.random.choice(min_frequency_names)
                        if roster_entry in members_not_assigned: 
                            assign_table.at[roster_entry,role] = 1
                            members_not_assigned = np.delete(members_not_assigned,np.where(members_not_assigned == roster_entry))
                            new_roster.at[slots_not_assigned[0],meeting_date] = roster_entry
                            role_pools_meeting_sp[role] = np.delete(role_pools_meeting_sp[role],np.where(role_pools_meeting_sp[role] == roster_entry))
                            if LOG:
                                print(meeting_date," ",slots_not_assigned[0],":",roster_entry," ","Second pass non-assigned member new slot",
                                      " ","Distance: ",distance_table.at[roster_entry,role],file=f) 
                            slots_not_assigned.remove(slots_not_assigned[0])
                        else:                    
                            old_role = assign_table.columns.values[np.where(assign_table.loc[roster_entry,:])]
                            assign_table.at[roster_entry,old_role[0]] = 0     #TODO: check if this works correctly, old_role should be an array
                            assign_table.at[roster_entry,role] = 1
                            old_slot_index = np.where(new_roster.loc[:,meeting_date] == roster_entry)
                            old_slot = new_roster.index.values[old_slot_index][0]
                            new_roster.at[old_slot,meeting_date] = ''
                            new_roster.at[slots_not_assigned[0],meeting_date] = roster_entry
                            role_pools_meeting_sp[role] = np.delete(role_pools_meeting_sp[role],np.where(role_pools_meeting_sp[role] == roster_entry))
                            if LOG:
                                print(meeting_date," ",slots_not_assigned[0],":",roster_entry," ","Second pass assigned member new slot"
                                      ," ","Distance: ",distance_table.at[roster_entry,role],file=f) 
                            slots_not_assigned.remove(slots_not_assigned[0])
                            slots_not_assigned.append(old_slot)
                            for m in members_not_assigned:
                                vacated_role = match_roster_role(old_slot)
                                if distance_table.at[m,vacated_role] > distance_thresholds[vacated_role] and m in role_pools_meeting_sp[vacated_role]:
                                    assign_table.at[m,vacated_role] = 1
                                    new_roster.at[old_slot,meeting_date] = m
                                    role_pools_meeting_sp[vacated_role] = np.delete(role_pools_meeting_sp[vacated_role],np.where(role_pools_meeting_sp[vacated_role] == m))
                                    if LOG:
                                        print(meeting_date," ",old_slot,":",m," ","Second pass non-assigned member old slot"
                                              ," ","Distance: ",distance_table.at[m,vacated_role],file=f) 
                                    members_not_assigned = np.delete(members_not_assigned,np.where(members_not_assigned == m))
                                    slots_not_assigned.remove(old_slot)
                                    break
            if LOG:
                print(meeting_date,"  Members not assigned after second stage: ",members_not_assigned, file=f)  
                print(meeting_date,"  Slots not assigned after second stage: ",slots_not_assigned, file=f)                     

        # THIRD PASS
        # Done if there are still unassinged slots after the first two passes. In this case, already assigned membeers will 
        # have to take multiple roles. Allowed pairs of roles are defined in the double_roles spreadsheet. Among the allowed 
        # combinations, weights_table determines the path of the least resistance. For example, Grammarian and Um & Ah counter
        # are more likely to be paired together than Grammarian and Evaluator.
                    
        if len(slots_not_assigned) > 0:
            slots_not_assigned_c = slots_not_assigned.copy()
            for slot in slots_not_assigned: 
                role = match_roster_role(slot)
                min_weight_names = get_weights(weights_table,assign_table, role_pools_meeting[role],role)
                for member in min_weight_names:
                    if (distance_table.at[member,role]) > distance_thresholds[role]:
                        assigned_roles = assign_table.columns.values[np.where(assign_table.loc[member,:])]    
                        assigned_roles = np.delete(assigned_roles,np.where(assigned_roles == "Number of roles"))
                        pair_allowed = 1
                        for role_a in assigned_roles: 
                            if double_roles[member_level[member]].at[role,role_a] == 'N':
                                pair_allowed = 0
                        if pair_allowed:
                            new_roster.at[slot,meeting_date] = member 
                            if LOG:
                                print(meeting_date," ",slot,":",member," ","Third pass"," ","Distance: ",distance_table.at[member,role],file=f) 
                            assign_table.at[member,role] = 1
                            assign_table.at[member,"Number of roles"] = assign_table.at[member,"Number of roles"] + 1
                            slots_not_assigned_c.remove(slot)
                            break
            if LOG:
                print(meeting_date,"  Members not assigned after third stage: ",members_not_assigned, file=f)  
                print(meeting_date,"  Slots not assigned after third stage: ",slots_not_assigned_c, file=f)

            if len(slots_not_assigned_c) > 0: 
                print("\n", "MEETING SLOTS NOT ASSIGNED ON: ",meeting_date)
                print(slots_not_assigned_c, "\n")             

        # NON MEETING ROLES

        mg_slots = ["Set Up/Pack Up #1","Set Up/Pack Up #2","Meet & Greet #1","Meet & Greet #2"]
        for slot in mg_slots:
            role = match_roster_role(slot)
            if len(role_pools_meeting[role]) > 0:
                [max_distance_names,above_threshold] = get_maximum_distance(distance_table,role_pools_meeting[role],role,distance_thresholds)
                if above_threshold:
                    if (len(max_distance_names) == 1):
                        roster_entry = max_distance_names[0]
                    else:
                        min_frequency_names = get_minimum_frequency(frequency_table,role,max_distance_names) 
                        if (len(min_frequency_names)) == 1: 
                            roster_entry = min_frequency_names[0]
                        else: 
                            if DEBUG: 
                                np.random.seed(0)
                            roster_entry = np.random.choice(min_frequency_names)
                    new_roster.at[slot,meeting_date] = roster_entry 
                    role_pools_meeting["Meet & Greet"] = np.delete(role_pools_meeting["Meet & Greet"],np.where(role_pools_meeting["Meet & Greet"] == roster_entry))
                    role_pools_meeting["Set Up/Pack Up"] = np.delete(role_pools_meeting["Set Up/Pack Up"],np.where(role_pools_meeting["Set Up/Pack Up"] == roster_entry))

        new_roster.at['Supper #1',meeting_date] = new_roster.at['Toastmaster',meeting_date]
        new_roster.at['Supper #2',meeting_date] = new_roster.at['Tabletopic Master',meeting_date]

        # Update distance, frequency, unassigned members

        distance_table.iloc[:,:] = distance_table.iloc[:,:] + 1
        for slot in new_roster.index.values:
            if (slot not in ["Business","Supper #1","Supper #2"]):
                member = new_roster.at[slot,meeting_date]
                if member in role_availability.index.values: 
                    role = match_roster_role(slot)
                    distance_table.at[member,role] = 1
                    frequency_table.at[member,role] = frequency_table.at[member,role] + 1 

        update_unassigned_meetings(unassigned_meetings,members_not_assigned)

        for slot in new_roster.index.values:
            member = new_roster.at[slot,meeting_date] 
            if member in role_availability.index.values: 
                if roles_assigned.at[member,meeting_date] == "":
                    roles_assigned.at[member,meeting_date] = slot  
                else: 
                    roles_assigned.at[member,meeting_date] = roles_assigned.at[member,meeting_date] + "," + slot

    print("\n",new_roster)
    writer = pd.ExcelWriter('new_roster.xlsx', engine='xlsxwriter')
    new_roster.to_excel(writer, sheet_name='Current Roster')
    roles_assigned.to_excel(writer, sheet_name='Members')
    writer.save()

if __name__ == '__main__':
    main()


