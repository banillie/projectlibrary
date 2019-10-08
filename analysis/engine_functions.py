'''Store of common functions found across the different programmes'''

def all_milestone_data_bulk(project_list, master_data):
    """Function to filter out ALL milestone data"""
    upper_dict = {}

    for name in project_list:
        try:
            p_data = master_data.data[name]
            lower_dict = {}
            for i in range(1, 50):
                try:
                    try:
                        lower_dict[p_data['Approval MM' + str(i)]] = \
                            {p_data['Approval MM' + str(i) + ' Forecast / Actual']: p_data[
                                'Approval MM' + str(i) + ' Notes']}
                    except KeyError:
                        lower_dict[p_data['Approval MM' + str(i)]] = \
                            {p_data['Approval MM' + str(i) + ' Forecast - Actual']: p_data[
                                'Approval MM' + str(i) + ' Notes']}

                    lower_dict[p_data['Assurance MM' + str(i)]] = \
                        {p_data['Assurance MM' + str(i) + ' Forecast - Actual']: p_data[
                                'Assurance MM' + str(i) + ' Notes']}
                except KeyError:
                    pass

            for i in range(18, 67):
                try:
                    lower_dict[p_data['Project MM' + str(i)]] = \
                        {p_data['Project MM' + str(i) + ' Forecast - Actual']: p_data['Project MM' + str(i) + ' Notes']}
                except KeyError:
                    pass
        except KeyError:
            lower_dict = {}

        upper_dict[name] = lower_dict

    return upper_dict

def ap_p_milestone_data_bulk(project_list, master_data):
    """Function to filter out approval and project delivery milestones"""
    upper_dict = {}

    for name in project_list:
        try:
            p_data = master_data.data[name]
            lower_dict = {}
            for i in range(1, 50):
                try:
                    try:
                        lower_dict[p_data['Approval MM' + str(i)]] = \
                            {p_data['Approval MM' + str(i) + ' Forecast / Actual'] : p_data['Approval MM' + str(i) + ' Notes']}
                    except KeyError:
                        lower_dict[p_data['Approval MM' + str(i)]] = \
                            {p_data['Approval MM' + str(i) + ' Forecast - Actual'] : p_data['Approval MM' + str(i) + ' Notes']}

                except KeyError:
                    pass

            for i in range(18, 67):
                try:
                    lower_dict[p_data['Project MM' + str(i)]] = \
                        {p_data['Project MM' + str(i) + ' Forecast - Actual'] : p_data['Project MM' + str(i) + ' Notes']}
                except KeyError:
                    pass
        except KeyError:
            lower_dict = {}

        upper_dict[name] = lower_dict

    return upper_dict

def assurance_milestone_data_bulk(project_list, master_data):
    """Function to filter out assurance milestone data"""
    upper_dict = {}

    for name in project_list:
        try:
            p_data = master_data.data[name]
            lower_dict = {}
            for i in range(1, 50):
                lower_dict[p_data['Assurance MM' + str(i)]] = \
                    {p_data['Assurance MM' + str(i) + ' Forecast - Actual']: p_data['Assurance MM' + str(i) + ' Notes']}

            upper_dict[name] = lower_dict
        except KeyError:
            upper_dict[name] = {}

    return upper_dict

def project_time_difference(proj_m_data_1, proj_m_data_2, date_of_interest):
    """Function that calculates time different between milestone dates"""
    upper_dict = {}

    for proj_name in proj_m_data_1:
        td_dict = {}
        for milestone in proj_m_data_1[proj_name]:
            if milestone is not None:
                milestone_date = tuple(proj_m_data_1[proj_name][milestone])[0]
                try:
                    if date_of_interest <= milestone_date:
                        try:
                            old_milestone_date = tuple(proj_m_data_2[proj_name][milestone])[0]
                            time_delta = (milestone_date - old_milestone_date).days  # time_delta calculated here
                            if time_delta == 0:
                                td_dict[milestone] = 0
                            else:
                                td_dict[milestone] = time_delta
                        except (KeyError, TypeError):
                            td_dict[milestone] = 'Not reported' # not reported that quarter
                except (KeyError, TypeError):
                    td_dict[milestone] = 'No date provided' # date has now been removed

        upper_dict[proj_name] = td_dict

    return upper_dict

def filter_group(dictionary, group_of_interest):
    """both below functions are in development"""
    project_list = []
    for project in dictionary:
        if dictionary[project]['DfT Group'] == group_of_interest:
            project_list.append(project)

    return project_list

def filter_gmpp(dictionary):
    project_list = []
    for project in dictionary:
        if dictionary[project]['GMPP - IPA ID Number'] is not None:
            project_list.append(project)

    return project_list

def bc_ref_stages(project_list, masters_list):
    """One of key functions used for calculating which quarter to baseline data from.

    Function returns a dictionary structured in the following way project name[('latest quarter info', 'latest bc'),
    ('last quarter info', 'last bc'), ('last baseline quarter info', 'last baseline bc'), ('oldest quarter info',
    'oldest bc')] depending on the amount information available in the data. Only the first three key values are returned,
    to ensure consistency (which is helpful later).

    project_list: list of project names
    masters_list = list of master dictionaries

    """
    output = {}

    for project_name in project_list:
        #print(name)
        all_list = []      # format [('quarter info': 'bc')] across all masters including project
        bl_list = []        # format ['bc', 'bc'] across all masters. bl_list_2 removes duplicates
        ref_list = []       # format as for all list but only contains the three tuples of interest
        for master in masters_list:
            try:
                bc_stage = master.data[project_name]['BICC approval point']
                quarter = master.data[project_name]['Reporting period (GMPP - Snapshot Date)']
                tuple = (quarter, bc_stage)
                all_list.append(tuple)
            except KeyError:
                pass

        for i in range(0, len(all_list)):
            bl_list.append(all_list[i][1])

        '''below lines of text from stackoverflow. Question, remove duplicates in python list while
        preserving order'''
        seen = set()
        seen_add = seen.add
        bl_list_2 = [x for x in bl_list if not (x in seen or seen_add(x))]

        ref_list.insert(0, all_list[0])     # puts the latest info into the list first

        try:
            ref_list.insert(1, all_list[1])    # puts that last info into the list
        except IndexError:
            ref_list.insert(1, all_list[0])

        if len(bl_list_2) == 1:                     # puts oldest info into list (as basline if no baseline)
            ref_list.insert(2, all_list[-1])
        else:
            for i in range(0, len(all_list)):      # puts in baseline
                if all_list[i][1] == bl_list[0]:
                    ref_list.insert(2, all_list[i])

        '''there is a hack here i.e. returning only first three in ref_list. There's a bug which I don't fully
        understand, but this solution is hopefully good enough for now'''
        output[project_name] = ref_list[0:3]

    return output

def get_master_baseline(project_list, masters_list, baselines_list):
    """
    Another key function used for calcualting which quarter to baseline data from.

    Fuction returns a dictionary structured in the following way project_name[n,n,n]. The n (number) values denote where
    the relevant quarter master dictionary is positions in the list of master dictionaries

    project_list: list of projects
    masters_list: list of masters
    baseline_list: list of project baseline information in the structure created by bc_ref_stage
    """
    output = {}

    for project_name in project_list:
        master_q_list = []
        for key in baselines_list[project_name]:
            for x, master in enumerate(masters_list):
                try:
                    quarter = master.data[project_name]['Reporting period (GMPP - Snapshot Date)']
                    if quarter == key[0]:
                        master_q_list.append(x)
                except KeyError:
                    pass

        output[project_name] = master_q_list

    return output

def convert_rag_text(dca_rating):

    if dca_rating == 'Green':
        return 'G'
    elif dca_rating == 'Amber/Green':
        return 'A/G'
    elif dca_rating == 'Amber':
        return 'A'
    elif dca_rating == 'Amber/Red':
        return 'A/R'
    elif dca_rating == 'Red':
        return 'R'
    else:
        return 'None'