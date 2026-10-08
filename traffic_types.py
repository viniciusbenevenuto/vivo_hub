"""traffic_types.py — Tabela oficial dos tipos de tráfego do PTI.

Fonte: planilha "No de A Formato de Entrega por Rota.xlsx"
(aba "Tipo de Rota e Formato_") — padrão de Interconexão da VIVO.

Cada tipo carrega as informações padrão que a aba Encaminhamento do PTI
preenche automaticamente: Formato do Nº de A, os dois sentidos de
encaminhamento, o CODEC e a sinalização.

Os marcadores do documento (Op_B, CN, XY) são mantidos LITERAIS, como no
padrão oficial. Para atualizar, edite este arquivo — nenhuma outra parte
do código precisa mudar.
"""
from typing import Final


TRAFFIC_TYPES: Final[tuple[dict[str, str], ...]] = (
    {
        "tipo":         'Tráf LC - CN XY (N8)',
        "formato_no_a": 'CN PREF-MCDU',
        "enc_ab":       'Origem VIVO-STFC chamando LC para a Op_B no Formato PREF-MCDU ou 9090 PREF-MCDU no mesmo CN. SE Municipal do CN - 15x / 19x . A VIVO NÃO envia nem recebe o RN3 (060)',
        "enc_ba":       'Origem Op_B chamando LC para a VIVO-STFC no mesmo CN no Formato PREF-MCDU ou 9090 PREF-MCDU. SE Municipal do CN - 15x / 19x, na Região III, conforme tabela de Tridígitos. A VIVO NÃO envia nem recebe o RN3 (060)',
        "codec":        'G-711 / G-729',
        "obs":          'SIP-I',
    },
    {
        "tipo":         'Tráf LC - CN XY (0CN)',
        "formato_no_a": 'CN PREF-MCDU',
        "enc_ab":       'Origem VIVO-STFC chamando LC para a Op_B no Formato 0CN PREF-MCDU ou 90CN PREF-MCDU no mesmo CN XY. SE Municipal do CN - 15x / 19x. A VIVO NÃO envia nem recebe o RN3 (060)',
        "enc_ba":       'Origem Op_B chamando LC para a VIVO-STFC no Formato 0CN PREF-MCDU ou 90CN PREF-MCDU no mesmo CN. SE Municipal do CN 15x / 19x, conforme tabela de Tridígitos na Região III. A VIVO NÃO envia nem recebe o RN3 (060)',
        "codec":        'G-711 / G-729',
        "obs":          'SIP-I',
    },
    {
        "tipo":         'Tráf LC + Trans LC - CN XY (N8)',
        "formato_no_a": 'CN PREF-MCDU',
        "enc_ab":       'Origem VIVO-STFC / Outras chamando LC para a Op_B no Formato PREF-MCDU ou 9090 PREF-MCDU no mesmo CN XY. SE Municipal do CN - 15x / 19x. A VIVO NÃO envia nem recebe o RN3 (060)',
        "enc_ba":       'Origem Op_B chamando LC para a VIVO-STFC / Outras no Formato PREF-MCDU ou 9090 PREF-MCDU no mesmo CN. SE Municipal do CN - 15x / 19x, conforme tabela de Tridígitos na Região III. A VIVO NÃO envia nem recebe o RN3 (060)',
        "codec":        'G-711 / G-729',
        "obs":          'SIP-I',
    },
    {
        "tipo":         'Tráf LC + Trans LC - CN XY (0CN)',
        "formato_no_a": 'CN PREF-MCDU',
        "enc_ab":       'Origem VIVO-STFC / Outras chamando LC para a Op_B no Formato 0CN PREF-MCDU ou 90CN PREF-MCDU no mesmo CN XY. SE Municipal do CN - 15x / 19x. A VIVO NÃO envia nem recebe o RN3 (060)',
        "enc_ba":       'Origem Op_B chamando LC para a VIVO-STFC / Outras no formato 0CN PREF-MCDU ou 90CN PREF-MCDU no mesmo CN SE Municipal do CN - 15x / 19x, conforme tabela de Tridígitos na Região III. A VIVO NÃO envia nem recebe o RN3 (060)',
        "codec":        'G-711 / G-729',
        "obs":          'SIP-I',
    },
    {
        "tipo":         'LD15 + CNG',
        "formato_no_a": 'LDN : CN PREF-MCDU\nLDI: E.164 : <No. Internacional>',
        "enc_ab":       'Origem VIVO-STFC ou Outra chamando com LD15 no formato 015 CN PREF-MCDU ou 9015 CN PREF MCDU para a Op_B Terminar da chamada Op_B.. A VIVO NÃO envia nem recebe o RN3 (060)',
        "enc_ba":       'Origem Op_B do CN chamando com LD15, no formato 015 CN PREF-MCDU, ou 9015 CN PREF-MCDU ou LDI 0015+No de B, para a VIVO-STFC realizar o encaminhamento da chamada. Origem Op_B do CN chamando CNG VIVO: 08XX 03XX 05XX 09XX (10-11 dig) - SE: 0CN 10315 / 0CN 10615. A VIVO NÃO envia nem recebe o RN3 (060)',
        "codec":        'G-711 / G-729',
        "obs":          'SIP-I',
    },
    {
        "tipo":         'LD XY + CNG Oper.',
        "formato_no_a": 'LDN : CN PREF-MCDU\nLDI: E.164 : <No. Internacional>',
        "enc_ab":       'Origem VIVO-STFC ou Outra do CN chamando com LD XY no formato 0XY CN PREF-MCDU ou 90XY CN PREF MCDU para a Op_B realizar o encaminhamento das chamadas. Origem VIVO-STFC ou Outra Chamando CNG Op_B: 08XX 03XX 05XX 09XX (10-11 dig) - SE: 103XY / 106XY A VIVO NÃO envia nem recebe o RN3 (060)',
        "enc_ba":       'Origem Op_B chamando com o LD XY no formato 0XY CN PREF-MCDU ou 90XY CN PREF-MCDU para a VIVO-STFC Terminar da chamada. A VIVO NÃO envia nem recebe o RN3 (060)',
        "codec":        'G-711 / G-729',
        "obs":          'SIP-I',
    },
    {
        "tipo":         'LD S/CSP + CNG Oper.',
        "formato_no_a": 'CN PREF-MCDU',
        "enc_ab":       'Origem VIVO-STFC Chamando para CNG Op_B 08XX 03XX 05XX 09XX (10-11 dig), da Op_B. A VIVO NÃO envia nem recebe o RN3 (060)',
        "enc_ba":       'Origem Op_B chamando com o LD s/CSP para a VIVO-STFC no formato 0CN PREF-MCDU ou 90CN PREF-MCDU para a Terminação da chamada. A VIVO NÃO envia nem recebe o RN3 (060)',
        "codec":        'G-711 / G-729',
        "obs":          'SIP-I',
    },
    {
        "tipo":         'TRANSP LD XY',
        "formato_no_a": 'CN PREF-MCDU',
        "enc_ab":       'N/A',
        "enc_ba":       'Origem Op_B chamando com o LD XY no formato 0XY CN PREF-MCDU para a VIVO-STFC Transportar a chamada. A VIVO NÃO envia nem recebe o RN3 (060)',
        "codec":        'G-711 / G-729',
        "obs":          'SIP-I',
    },
    {
        "tipo":         'TRANSP LD s/CSP',
        "formato_no_a": 'CN PREF-MCDU',
        "enc_ab":       'N/A',
        "enc_ba":       'Origem Op_B chamando com o LD s/CSP no formato 0CN PREF-MCDU para a VIVO-STFC Transportar da chamada. A VIVO NÃO envia nem recebe o RN3 (060)',
        "codec":        'G-711 / G-729',
        "obs":          'SIP-I',
    },
    {
        "tipo":         'VC1 - CN',
        "formato_no_a": 'CN PREF-MCDU',
        "enc_ab":       'Origem VIVO-SMP chamando LC no Formato PREF-MCDU ou 9090 PREF-MCDU para a Op_B no mesmo CN. A VIVO NÃO envia nem recebe o RN3 (060)',
        "enc_ba":       'Origem Op_B-STFC chamando LC no formato 0CN PREF-MCDU ou 90CN PREF-MCDU para a VIVO-SMP no mesmo CN. SE 1058 / 0CN 1058. A VIVO NÃO envia nem recebe o RN3 (060)',
        "codec":        'G-711 / AMR',
        "obs":          'SIP-I',
    },
    {
        "tipo":         'VC1 - CN (EIR)',
        "formato_no_a": 'CN PREF-MCDU',
        "enc_ab":       'Origem VIVO-SMP chamando LC no Formato 0CN PREF-MCDU ou 90CN PREF-MCDU para a Op_B no mesmo CN. A VIVO NÃO envia nem recebe o RN3 (060)',
        "enc_ba":       'Origem Op_B-STFC chamando LC no formato 0CN PREF-MCDU ou 90CN PREF-MCDU para a VIVO-SMP no mesmo CN. SE 1058 / 0CN 1058. A VIVO NÃO envia nem recebe o RN3 (060)',
        "codec":        'G-711 / AMR',
        "obs":          'SIP-I',
    },
    {
        "tipo":         'VC2 / VC3 - CN XY - LD XY',
        "formato_no_a": 'LDN : CN PREF-MCDU\nLDI: E.164 : <No. Internacional>',
        "enc_ab":       'Origem VIVO-SMP chamando LD no Formato 0XY CN PREF-MCDU ou 90XY CN PREF-MCDU para a Op_B em CNs diferentes. CNG Op_B 08XX 03XX 05XX 09XX (10-11 dig), da Op_B A VIVO NÃO envia nem recebe o RN3 (060)',
        "enc_ba":       'Origem Op_B-STFC chamando LDXY no formato 0XY CN PREF-MCDU ou 90XY CN PREF-MCDU em CN diferentes para a VIVO-SMP terminar a chamada. A VIVO NÃO envia nem recebe o RN3 (060)',
        "codec":        'G-711 / AMR',
        "obs":          'SIP-I',
    },
    {
        "tipo":         'VC2 / VC3 - CN XY LD s/CSP',
        "formato_no_a": 'CN PREF-MCDU',
        "enc_ab":       'Origem VIVO-SMP chamando CNG Op_B 08XX 03XX 05XX 09XX (10-11 dig), da Op_B',
        "enc_ba":       'Origem Op_B-STFC chamando LDXY no formato 0CN PREF-MCDU ou 9 0CN PREF-MCDU em CN diferentes para a VIVO-SMP terminar a chamada. A VIVO NÃO envia nem recebe o RN3 (060)',
        "codec":        'G-711 / AMR',
        "obs":          'SIP-I',
    },
    {
        "tipo":         'STIR SHAKEN (TRANSP LD XY)',
        "formato_no_a": 'CN PREF-MCDU',
        "enc_ab":       'N/A',
        "enc_ba":       'Origem Op_B chamando LD no no formato 0XY CN PREF-MCDU para a VIVO-SMP / TIM-SMP / CLARO-SMP para a VIVO Transportar a chamada. A VIVO NÃO envia nem recebe o RN3 (060)',
        "codec":        'G-711',
        "obs":          'SIP-I',
    },
    {
        "tipo":         'STIR SHAKEN (TRANSP LD s/CSP)',
        "formato_no_a": 'CN PREF-MCDU',
        "enc_ab":       'N/A',
        "enc_ba":       'Origem Op_B chamando LD no no formato 0CN PREF-MCDU para a VIVO-SMP / TIM-SMP / CLARO-SMP para a VIVO Transportar a chamada. A VIVO NÃO envia nem recebe o RN3 (060)',
        "codec":        'G-711',
        "obs":          'SIP-I',
    },
)

# Nome do tipo -> registro completo
TRAFFIC_BY_NAME: Final[dict[str, dict[str, str]]] = {
    t["tipo"]: t for t in TRAFFIC_TYPES
}

# Ordem oficial dos tipos (usada nos chips do formulário)
TRAFFIC_TYPE_NAMES: Final[tuple[str, ...]] = tuple(
    t["tipo"] for t in TRAFFIC_TYPES
)
