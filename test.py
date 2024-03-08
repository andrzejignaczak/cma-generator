try:
    import cma
    import tkinter as tk
    import pandas as pd
    from tkinter import filedialog 


    def find_mold_variables(search_value):
        df = pd.read_excel('PL03RobotPrograms.xlsx')
        index = df[df['KOD FORMY'] == search_value].index
        
        index = index[0]
        m_length = df.iloc[index, 2]
        m_width = df.iloc[index, 3]
        m_drain_type = df.iloc[index, 4]
        m_C1 = df.iloc[index, 5]
        m_C2 = df.iloc[index, 6]
        return {
            "m_length": m_length,
            "m_width": m_width,
            "m_drain_type": m_drain_type,
            "m_C1": m_C1,
            "m_C2": m_C2
        }

    # obj_name = open_file_dialog()
    obj_name = "G0002.obj"
    # search_value = input("Mould Number: ")
    search_value = "G0002"
    variables = find_mold_variables(search_value)

    # Print the values of m_length and m_width
    print("m_length:", variables["m_length"])
    print("m_width:", variables["m_width"])

    # Assuming cma and related objects are already defined

    program_path = cma.script_folder + "\\Default_Program.csv"
    cma.act_prg = cma.read_program(program_path)
    cma.act_prg.model_path = obj_name
    a_length = variables["m_length"]
    a_width = variables["m_width"]
    print("a_width:", a_width, "a_length:", a_length)  # corrected print statement
    iy = float(a_length)
    ix = float(a_width)
    x = ix - 40
    y = iy - 40



    for active_component in cma.act_prg.cmp:
        for instruction_active in active_component.ptp:
            if instruction_active.q_c[0] >= 0:
                instruction_active.q_c[0] = (x/2) + instruction_active.q_c[0]
            else: instruction_active.q_c[0] = (-x/2) -(-instruction_active.q_c[0])
            if instruction_active.q_c[1] >= 0:
                instruction_active.q_c[1] = (y/2) + instruction_active.q_c[1]
            else: instruction_active.q_c[1] = (-y/2) -(-instruction_active.q_c[1])
    
    cma.done()
except Exception as e:
	print_error(e)
# input()
