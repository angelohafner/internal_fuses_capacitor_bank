import numpy as np

def star_voltages(Vabc, Zabco):
    # Unpack line voltages and impedances
    Vab, Vbc, Vca = Vabc
    Zao, Zbo, Zco = Zabco

    Van, Vbn, Vcn = np.exp(-1j*np.pi/6) / np.sqrt(3) * np.array(Vabc)
    Vabcn = np.array([Van, Vbn, Vcn])

    # Calculate reference phase voltages (assuming zero neutral displacement)
    matriz_v = np.array([[Vab],
                         [Vbc]], dtype=np.complex128)

    matriz_z = np.array([[Zao + Zbo, -Zbo],
                         [-Zbo, Zbo + Zco]], dtype=np.complex128)

    matriz_i = np.linalg.inv(matriz_z) @ matriz_v

    Iao = matriz_i[0]
    Ibo = matriz_i[1] - matriz_i[0]
    Ico = -matriz_i[1]

    Vao = Iao * Zao
    Vbo = Ibo * Zbo
    Vco = Ico * Zco

    Iabco = np.array([Iao, Ibo, Ico])
    Vabco = np.array([Vao, Vbo, Vco])

    Von_a = Van - Vao
    Von_b = Van - Vao
    Von_c = Van - Vao
    
    if np.abs(Von_a - Von_b) > 0.001 or \
       np.abs(Von_b - Von_c) > 0.001 or \
       np.abs(Von_c - Von_a) > 0.001:    
        print("algo deu errado!")

    # fig = phasor_with_neutral([Vao, Vbo, Vco, Von], Vabcn, Vabc, plot_title='Diagrama Fasorial Trifásico')

    return Vabcn, Vabco, Von_a, Iabco
