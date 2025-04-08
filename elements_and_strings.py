import numpy as np
from capacitor_equivalente import equivalents_from_matrix


def compute_elements_units_strings(X, Su, N, P, Pa, Pt, S, index):
    # create base matrix of elements
    elements_matrix = 1j * X * np.ones((Su, N))
    elements_matrix_bad = elements_matrix.copy()
    elements_matrix_bad[0, 0:index] = 1e+99 *elements_matrix_bad[0, 0:index]

    # this in only to confirm ieee 37.99
    parallels_internal_group_bad = 1 / np.sum( 1/elements_matrix_bad[0, :])
    parallels_internal_group = 1 / np.sum( 1/elements_matrix[0, :])
    Ci = parallels_internal_group / parallels_internal_group_bad

    # compute cell equivalents
    equiv_cell, parallels_units = equivalents_from_matrix(elements_matrix)
    equiv_cell_bad, parallels_units_bad = equivalents_from_matrix(elements_matrix_bad)

    # this in only to confirm ieee 37.99
    Cu = equiv_cell / equiv_cell_bad

    # build cell sets for strings
    cells_set = equiv_cell * np.ones((1, P))
    cells_set_bad = cells_set.copy()
    cells_set_bad[0, 0] = equiv_cell_bad

    # cells_set_bad = equiv_cell_bad * np.ones((1, P))
    Pst = Pt - Pa - P
    cells_set_second_type = equiv_cell * np.ones((1, Pst))

    # compute set equivalents (still need put in series for string)
    equiv_set, parallels_set = equivalents_from_matrix(cells_set)
    equiv_set_bad, parallels_set_bad = equivalents_from_matrix(cells_set_bad)
    equiv_set_second_type, parallels_set_second_type = \
        equivalents_from_matrix(cells_set_second_type)

    # this in only to confirm ieee 37.99
    Cg = equiv_set / equiv_set_bad

    # compute string equivalents
    equiv_string = S * equiv_set
    equiv_string_bad = (S - 1) * equiv_set + equiv_set_bad
    equiv_string_second_type = S * equiv_set_second_type

    # this in only to confirm ieee 37.99
    Cs = equiv_string / equiv_string_bad

    return (
        equiv_cell, parallels_units,
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
        Cs
    )



def analyze_branches_and_phases(equiv_string, equiv_string_bad, equiv_string_second_type,
                                Pa, P, Vabc,
                                equivalents_from_matrix, parallel_impedance,
                                star_voltages):
    # number of parallel strings per phase
    Ps = int(Pa / P)

    # left branch with good strings
    branch_matrix_left = equiv_string * np.ones((1, Ps))

    # left branch with one bad string
    branch_matrix_left_bad = branch_matrix_left.copy()
    branch_matrix_left_bad[0, 0] = equiv_string_bad

    # right branch with two different string types
    branch_matrix_right = np.array([[equiv_string, equiv_string_second_type]])

    # equivalent impedances for branches
    equiv_phase_left, parallels_phase_left = equivalents_from_matrix(branch_matrix_left)
    equiv_phase_right, parallels_phase_second_right = equivalents_from_matrix(branch_matrix_right)
    equiv_phase_left_bad, parallels_phase_left_bad = equivalents_from_matrix(branch_matrix_left_bad)

    # phase impedances (bad string in phase A only)
    phase_a = parallel_impedance(equiv_phase_left_bad, equiv_phase_right)
    phase_b = parallel_impedance(equiv_phase_left, equiv_phase_right)
    phase_c = parallel_impedance(equiv_phase_left, equiv_phase_right)

    # this in only to confirm ieee 37.99
    Cp = phase_b / phase_a

    Zabco = np.array([phase_a, phase_b, phase_c])

    # compute voltages and currents
    Vabcn, Vabco, Von, Iabco = star_voltages(Vabc, Zabco)

    return (
        branch_matrix_left, branch_matrix_left_bad, branch_matrix_right,
        equiv_phase_left, parallels_phase_left,
        equiv_phase_right, parallels_phase_second_right,
        equiv_phase_left_bad, parallels_phase_left_bad,
        phase_a, phase_b, phase_c,
        Zabco, Vabcn, Vabco, Von, Iabco,
        Cp
    )


def compute_branch_currents_and_voltages(Vabcn, Vabco,
                                         equiv_phase_left,
                                         equiv_phase_left_bad,
                                         equiv_phase_right,
                                         branch_matrix_left_bad,
                                         equiv_string_bad,
                                         equiv_set_bad,equiv_set,
                                         parallels_set_bad, parallels_set,
                                         equiv_cell_bad, parallels_units_bad,
                                         parallels_internal_group_bad):
    # branch currents (left and right)
    current_branch_left_a = Vabco[0] / equiv_phase_left_bad
    current_branch_right_a = Vabco[0] / equiv_phase_right

    current_branch_left_b = Vabco[1] / equiv_phase_left
    current_branch_right_b = Vabcn[1] / equiv_phase_right

    current_branch_left_c = Vabco[2] / equiv_phase_left
    current_branch_right_c = Vabco[2] / equiv_phase_right

    current_sum_branch_left = np.round(current_branch_left_a +
                                       current_branch_left_b +
                                       current_branch_left_c)

    current_sum_branch_right = np.round(current_branch_right_a +
                                        current_branch_right_b +
                                        current_branch_right_c)

    # voltage in bad branch (in equivalente branch)
    voltage_branch_left_bad = Vabco[0]

    # current in bad string
    current_equiv_string_bad = voltage_branch_left_bad / equiv_string_bad

    # voltage on set with bad unit
    Vcu = current_equiv_string_bad * equiv_set_bad

    # current in bad unit
    current_bad_unit = Vcu / equiv_cell_bad

    # voltages in elements of bad unit
    voltages_parallels_units_bad = parallels_units_bad * current_bad_unit

    # voltages on intenal group of bad cell
    # this in only to confirm ieee 37.99 #nao está batendo 100% ainda
    voltage_parallels_internal_group_bad = parallels_internal_group_bad * current_bad_unit


    return (
        current_branch_left_a, current_branch_right_a,
        current_branch_left_b, current_branch_right_b,
        current_branch_left_c, current_branch_right_c,
        current_sum_branch_left, current_sum_branch_right,
        voltage_branch_left_bad, current_equiv_string_bad,
        Vcu, current_bad_unit, voltages_parallels_units_bad,
        voltage_parallels_internal_group_bad
    )
