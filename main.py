import numpy as np
import pandas as pd
from exporta_para_excel import export_to_excel_with_formatting
import streamlit as st
from io import BytesIO
from capacitor_equivalente import equivalents_from_matrix, parallel_impedance
from star_voltages_and_currents import star_voltages
from elements_and_strings import compute_elements_units_strings, analyze_branches_and_phases, compute_branch_currents_and_voltages
# Configura o tema claro
st.set_page_config(
    page_title="Internal Fused Capacitor Bank",
    layout="centered",  # ou "wide"
    initial_sidebar_state="auto",
)

# CSS para ocultar o menu de troca de tema
st.markdown(
    """
    <style>
    [data-testid="stHeader"] button {
        display: none;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# 1) input_data
st.title("Internal Fused Capacitor Banks Unbalance Protection")

st.write("""
Este aplicativo apresenta uma tabela com variáveis relacionadas a cálculos de capacitores em sistemas de potência, 
incluindo suas fórmulas e descrições.
""")

st.image("Figure 34 Illustration of an uneven double wye-connected bank.png",
         caption="Illustration of an uneven double wye-connected bank")
# Create 3 columns
col1, col2, col3 = st.columns(3)

with col1:
    power_capacitor_bank_MVAr = st.number_input(
        "Power Capacitor Bank (MVAr)",
        min_value=0.0,
        value=10.0,
        step=0.1,
        format="%.1f"
    )

    voltage_line_kV = st.number_input(
        "Voltage Line (kV)",
        min_value=0.0,
        value=34.5,
        step=0.1,
        format="%.1f"
    )

    frequency_Hz = st.number_input(
        "Frequency (Hz)",
        min_value=0.0,
        value=60.0,
        step=0.1,
        format="%.1f"
    )

with col2:
    N = st.number_input(
        "N",
        min_value=0,
        value=14,
        step=1
    )

    Su = st.number_input(
        "Su",
        min_value=0,
        value=3,
        step=1
    )

    P = st.number_input(
        "P",
        min_value=0,
        value=3,
        step=1
    )

with col3:
    S = st.number_input(
        "S",
        min_value=0,
        value=4,
        step=1
    )

    Pa = st.number_input(
        "Pa",
        min_value=0,
        value=6,
        step=1
    )

    Pt = st.number_input(
        "Pt",
        min_value=0,
        value=11,
        step=1
    )


Vfase = 1e3*voltage_line_kV / np.sqrt(3)
a = np.exp(1j * 2 * np.pi / 3)
trinta = np.exp(1j * np.pi / 6)
Vabc = Vfase * np.sqrt(3) * np.array([trinta, trinta*(a**2), trinta*a])

w = 2*np.pi*frequency_Hz

Pst = Pt - Pa - P

corrente_fase = (1e6*power_capacitor_bank_MVAr/3) / Vfase
reatancia_banco = Vfase / corrente_fase
reatancia_celula = reatancia_banco * Pt / S
reatancia_elemento =  reatancia_celula * N / Su
X = reatancia_elemento

# Dados da tabela
data = [
    {
        "Variável": "Internal group per-unit capacitance (Ci)",
        "Fórmula": r"C_i = \frac{N}{f}",
        "Descrição": "Representa a capacitância por unidade (per-unit) de um grupo interno de capacitores, calculada com base no número de fusíveis (N) e no fator f (provavelmente relacionado à configuração ou frequência). É uma medida normalizada da capacitância do grupo."
    },
    {
        "Variável": "Internal group voltage (for capacitor unit at 1 per-unit voltage) (VG)",
        "Fórmula": r"V_g = \frac{(S_u \times N)}{[(S_u - 1) \times (N - f) + N]}",
        "Descrição": "É a tensão (em pu) que ocorre nos elementos de um grupo de capacitores quando a tensão total na unidade é 1 per-unit. Leva em conta o número de fusíveis (f), o número total de elementos (N) e a configuração (Su). A capacitância do grupo afetado é Ci."
    },
    {
        "Variável": "Capacitor unit per-unit capacitance (Cu)",
        "Fórmula": r"C_u = \frac{(S_u \times Ci)}{[Ci \times (S_u - 1) + 1]}",
        "Descrição": "Representa a capacitância por unidade (per-unit) de uma unidade de capacitor, considerando todos os grupos, exceto o grupo afetado, que tem capacitância Ci. Su é o número de grupos em série."
    },
    {
        "Variável": "Parallel group per-unit capacitance (Cg)",
        "Fórmula": r"C_g = \frac{P-1+C_u}{P}",
        "Descrição": "É a capacitância por unidade de um grupo de capacitores em paralelo, onde P é o número de unidades em paralelo e R é um fator relacionado à configuração do grupo. Para o grupo afetado, a capacitância por unidade é 1."
    },
    {
        "Variável": "Affected string capacitance (Cs)",
        "Fórmula": r"C_s = \frac{C_g}{(S \times [1 - 1])}",
        "Descrição": "Representa a capacitância por unidade de uma string (cadeia) de grupos de capacitores em paralelo afetada, onde CG é a capacitância do grupo e S é o número total de strings. Para o grupo afetado, a capacitância é CG; para os outros, é 1."
    },
    {
        "Variável": "Per-unit capacitance, phase with affected unit (Cp)",
        "Fórmula": r"C_p = \frac{(C_s \times P)}{(S - P)}",
        "Descrição": "É a capacitância por unidade de uma fase que contém a unidade afetada. CS é a capacitância da string afetada, P é o número de strings paralelas afetadas, e S é o número total de strings. Para as outras fases, a capacitância é 1."
    },
    {
        "Variável": "Neutral-to-ground per-unit of Vg (Vng)",
        "Fórmula": r"V_{ng} = (V_g \times [G - 1])",
        "Descrição": "Representa a tensão (em pu) entre o neutro e o terra para bancos de capacitores aterrados (G = 0) ou não aterrados (G = 1). Para bancos não aterrados, essa tensão é sempre 0. Para bancos aterrados, é a tensão do grupo afetado (VG) vezes (G - 1)."
    },
    {
        "Variável": "Voltage on affected phase (Vph)",
        "Fórmula": r"V_{ph} = V_{ng} + 1",
        "Descrição": "É a tensão (em pu) na fase afetada, considerando a tensão entre neutro e terra (VNG). Para unidades com fusíveis, a operação da fase afetada reduz a capacitância e aumenta a tensão, que é maior que 1."
    },
    {
        "Variável": "Voltage on affected unit (Vcu)",
        "Fórmula": r"V_{cu} = \frac{V_{ln} \times V_{cu}}{C_g}",
        "Descrição": "Representa a tensão (em pu) na unidade de capacitor afetada, com base na tensão da fase afetada (VPH) e no número de strings (S). Se CG = 0, a tensão é diretamente proporcional a VPH."
    },
    {
        "Variável": "Voltage on affected elements (Ve)",
        "Fórmula": r"V_e = V_{cu} \times V_g",
        "Descrição": "É a tensão (em pu) nos elementos afetados da unidade, calculada como o produto da tensão na unidade afetada (VCU) e a tensão do grupo (VG)."
    },
    {
        "Variável": "Current through affected capacitor(s) (Iu)",
        "Fórmula": r"I_u = V_{cu} \times C_u",
        "Descrição": "Representa a corrente (em pu) que passa pela unidade de capacitor afetada, proporcional à tensão (VCU) e à capacitância (CU). É usada para estimar o tempo de queima de fusíveis."
    },
    {
        "Variável": "Current in affected string (IST)",
        "Fórmula": r"I_{st} = C_s \times V_{ph}",
        "Descrição": "É a corrente (em pu) na string afetada, proporcional à capacitância da string (CS) e à tensão da fase (VPH). Pode ser usada para comparar correntes em diferentes esquemas."
    },
    {
        "Variável": "Current in affected phase (Iph)",
        "Fórmula": r"I_{ph} = C_p \times V_{ph}",
        "Descrição": "Representa a corrente (em pu) na fase afetada, calculada como o produto da capacitância da fase (CP) e da tensão da fase (VPH). Útil para ajustes de proteção."
    },
    {
        "Variável": "Ground current change (Ig)",
        "Fórmula": r"I_g = (1 - G) \times (1 - I_{ph})",
        "Descrição": "É a variação da corrente de terra, usada em esquemas de proteção que utilizam corrente entre neutro e terra ou tensão neutro-terra. G indica se o sistema é aterrado (G = 0) ou não (G = 1), e IPH é a corrente na fase afetada."
    },
    {
        "Variável": "Neutral current for ungrounded Y-Y banks (In)",
        "Fórmula": r"I_n = \frac{(3 \times V_{ng} \times G \times (P_t - P_a))}{P_t}",
        "Descrição": "Representa a corrente de neutro em bancos Y-Y não aterrados, calculada com base na tensão neutro-terra (VNG), no número de strings paralelas (P), e na fase (PH). É a corrente no neutro do banco não afetado."
    }
]

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

st.dataframe(df)

excel_file = 'unbalance_capacitor_back_protection.xlsx'
export_to_excel_with_formatting(
    df=df,
    file_path=excel_file,
    sheet_name='internal_fused',
    header_color='#4472C4',  # Azul corporativo
    even_row_color='#EAEAEA',  # Cinza claro
    odd_row_color='#FFFFFF',  # Branco
    border_color='#4472C4'  # Azul corporativo
)

st.download_button(
    label="Download Excel File",
    data=excel_file,
    file_name=excel_file,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Exibindo a tabela
st.subheader("Tabela de Variáveis")
for item in data:
    st.write(f"**{item['Variável']}**")
    st.latex(item['Fórmula'])
    st.write(item['Descrição'])
    st.write("---")