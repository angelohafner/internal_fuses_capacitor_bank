import numpy as np
import pandas as pd
from exporta_para_excel import export_to_excel_with_formatting
from input_data import Vfase, Vabc, X, N, Su, P, S, Pa, Pt
from capacitor_equivalente import equivalents_from_matrix, parallel_impedance
from star_voltages_and_currents import star_voltages
from elements_and_strings import compute_elements_units_strings, analyze_branches_and_phases, compute_branch_currents_and_voltages


# 1) input_data
# from input_data import .....

nnn:int = N
f__vector = 999 * np.ones(nnn)
# affected capacitances
Ci_vector = 999 * np.ones(nnn)
Cu_vector = 999 * np.ones(nnn)
Cg_vector = 999 * np.ones(nnn)
Cs_vector = 999 * np.ones(nnn)
Cp_vector = 999 * np.ones(nnn)
# voltages
Vph_vector = 999 * np.ones(nnn)
Vg__vector = 999 * np.ones(nnn) # first internal group affected cell
Vgn_vector = 999 * np.ones(nnn)
Vln_vector = 999 * np.ones(nnn)
Vcu_vector = 999 * np.ones(nnn)
Ve__vector = 999 * np.ones(nnn)
# currents
Iu__vector = 999 * np.ones(nnn)
Ist_vector = 999 * np.ones(nnn)
Iph_vector = 999 * np.ones(nnn)
Id__vector = 999 * np.ones(nnn)
Ist_vector = 999 * np.ones(nnn)
Iph_vector = 999 * np.ones(nnn)
In__vector = 999 * np.ones(nnn)

for index_bad in range(nnn):

    # 2) compute_elements_units_strings
    (equiv_cell, parallels_units,
     equiv_cell_bad, parallels_units_bad,
     cells_set, cells_set_bad, cells_set_second_type,
     equiv_set, parallels_set,
     equiv_set_bad, parallels_set_bad,
     equiv_set_second_type, parallels_set_second_type,
     equiv_string, equiv_string_bad, equiv_string_second_type,
     Ci,
     parallels_internal_group_bad,
     Cu,
     Cg,
     Cs) = \
        compute_elements_units_strings(X, Su, N, P, Pa, Pt, S, index_bad)


    # 3) analyze_branches_and_phases

    (branch_matrix_left, branch_matrix_left_bad, branch_matrix_right,
     equiv_phase_left, parallels_phase_left,
     equiv_phase_right, parallels_phase_second_right,
     equiv_phase_left_bad, parallels_phase_left_bad,
     phase_a, phase_b, phase_c,
     Zabco, Vabcn, Vabco, Von, Iabco,
     Cp) = \
        analyze_branches_and_phases(equiv_string,
                                    equiv_string_bad,
                                    equiv_string_second_type,
                                    Pa, P, Vabc,
                                    equivalents_from_matrix,
                                    parallel_impedance,
                                    star_voltages)


    # 4) compute_branch_currents_and_voltages

    (current_branch_left_a, current_branch_right_a,
     current_branch_left_b, current_branch_right_b,
     current_branch_left_c, current_branch_right_c,
     current_sum_branch_left, current_sum_branch_right,
     voltage_branch_left_bad, current_equiv_string_bad,
     Vcu, current_bad_unit, voltages_parallels_units_bad,
     voltage_parallels_internal_group_bad,
     current_sum_branch_left, current_sum_branch_right) = \
        compute_branch_currents_and_voltages(Vabcn, Vabco,
                                                 equiv_phase_left,
                                                 equiv_phase_left_bad,
                                                 equiv_phase_right,
                                                 branch_matrix_left_bad,
                                                 equiv_string_bad,
                                                 equiv_set_bad,equiv_set,
                                                 parallels_set_bad, parallels_set,
                                                 equiv_cell_bad, parallels_units_bad,
                                                 parallels_internal_group_bad)

    f__vector[index_bad] = index_bad

    Ci_vector[index_bad] = np.abs(Ci)  # internal group
    Cu_vector[index_bad] = np.abs(Cu)  # units
    Cg_vector[index_bad] = np.abs(Cg)  # groups of units
    Cs_vector[index_bad] = np.abs(Cs)  # string
    Cp_vector[index_bad] = np.abs(Cp)  # phase

    Vph_vector[index_bad] = np.abs(Vabco[0].item())
    Vgn_vector[index_bad] = np.abs(Von.item())
    Vln_vector[index_bad] = np.abs(Vabco[0].item())
    Vcu_vector[index_bad] = np.abs(Vcu.item() / (Vfase/S) )
    Ve__vector[index_bad] = np.abs(voltage_parallels_internal_group_bad.item() / (Vfase/(S*Su)))
    
    Iu__vector[index_bad] = np.abs(current_bad_unit.item())
    Ist_vector[index_bad] = np.abs(current_equiv_string_bad.item())
    Iph_vector[index_bad] = np.abs(Iabco[0].item())
    In__vector[index_bad] = np.abs(current_sum_branch_left.item())
    


dec = 8
df = pd.DataFrame({
    'burned': f__vector,
    'Ci' : np.round(Ci_vector, dec),
    'Cu' : np.round(Cu_vector, dec),
    'Cg' : np.round(Cg_vector, dec),
    'Cs' : np.round(Cs_vector, dec),
    'Cp' : np.round(Cp_vector, dec),
    'Vgn': np.round(Vgn_vector / Vfase, dec),
    'Vln': np.round(Vln_vector / Vfase, dec),
    'Vcu': np.round(Vcu_vector, dec),
    'Ve' : np.round(Ve__vector, dec),
    'Iu' : np.round(Iu__vector / Iu__vector[0], dec),
    'Ist': np.round(Ist_vector / Ist_vector[0], dec),
    'Iph': np.round(Iph_vector / Iph_vector[0], dec),
    'In' : np.round(In__vector / Iph_vector[0], dec),
    'Iph [A]' : np.round(Iph_vector, dec),
    'Vph [V]' : np.round(Vph_vector, dec),
    'In [A]'  : np.round(In__vector, dec),
    'Vgn [V]' : np.round(Vgn_vector, dec)
})

export_to_excel_with_formatting(
    df=df,
    file_path='unbalance_capacitor_back_protection.xlsx',
    sheet_name='internal_fused',
    header_color='#4472C4',  # Azul corporativo
    even_row_color='#EAEAEA',  # Cinza claro
    odd_row_color='#FFFFFF',  # Branco
    border_color='#4472C4'  # Azul corporativo
)
