def get_courses(prefix):
    fname = "c:\\temp\\pizza_hunt_graded\\enrollment.txt"
    prev_subj = ""
    prev_num = ""
    fvar = open(fname,"r")
    courses = {}
    count = 0
    for line in fvar:
        if count > 2: 
            line = line.strip()
            parts = line.split("\t")
            crn            = parts[0].strip()
            subj           = parts[1].strip()
            if subj == "":
                subj = prev_subj
            prev_subj = subj
            num            = parts[2].strip()
            if num == "":
                num = prev_num
            prev_num = num
            sec            = parts[3].strip()
            key = "%s-%s-%s-%s" % (prefix,subj,num,sec)
            title          = parts[4].strip()
            level          = parts[5].strip()   # for example, UG
            part_of_term   = parts[6].strip()   # for example, full term
            instructor     = parts[7].strip()
            campus         = parts[8].strip()   # for example, ROM
            cap_str        = parts[9].strip()
            if cap_str == "":
                capacity = 0
            else:
                capacity = int(cap_str)
            enrol_str      = parts[10].strip()
            if enrol_str == "":
                enrollment = 0
            else:
                enrollment = int(enrol_str)
            building       = parts[12].strip()
            room           = parts[13].strip()
            begin_time     = parts[14].strip()
            end_time       = parts[15].strip()
            days           = parts[16].strip()+parts[17].strip()+parts[18].strip()+parts[19].strip()+parts[20].strip()+parts[21].strip()
            college        = parts[23].strip()
            dept           = parts[24].strip()
            dept_desc      = parts[25].strip()
            status         = parts[26].strip()
            schedule       = parts[27].strip()
            sched_desc     = parts[28].strip()
            instr_method   = parts[29].strip()
            inst_meth_desc = parts[30].strip()
            course_info = {"crn":crn,"subj":subj,"num":num,"sec":sec,"title":title,"level":level,"part_of_term":part_of_term,"instructor":instructor,"campus":campus,
                           "capacity":capacity,"enrollment":enrollment,"building":building,"room":room,"begin_time":begin_time,"end_time":end_time,"days":days,
                           "college":college,"department":dept,"department_description":dept_desc,"status":status,"schedule":schedule,"schedule_description":sched_desc,
                           "instruction_method":instr_method,"instruction_method_description":inst_meth_desc}
            courses[key] = course_info
        count = count + 1
    return courses

courses = get_courses("SP19")
for key in courses:
    print("%s\t%s"%(key,courses[key]["title"]))

