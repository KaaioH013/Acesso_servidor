import pandas as pd

from src.conexao import get_engine


def main():
    engine = get_engine()

    q_desenho = """
    SELECT TOP 5
        d.CODIGO AS Cod_Desenho,
        d.DESENHO AS Nr_Desenho,
        d.DESCRICAO,
        d.DTCADASTRO,
        d.DTATUALIZADO,
        d.STATUS
    FROM PR_DESENHO d
    WHERE d.DESENHO = '3809029008'
    ORDER BY d.CODIGO DESC
    """

    q_lista_raiz = """
    SELECT TOP 10
        le.CODIGO AS ListaEng_Codigo,
        le.CODDESENHO,
        le.POSICAO,
        le.DTATUALIZADO
    FROM PR_LISTAENG le
    JOIN PR_DESENHO d ON d.CODIGO = le.CODDESENHO
    WHERE d.DESENHO = '3809029008'
      AND le.POSICAO = '00'
    ORDER BY le.CODIGO DESC
    """

    q_roteiro_opcao = """
    SELECT TOP 20
        o.CODIGO AS RoteiroOpcao,
        o.LISTAENG,
        o.CKPADRAO,
        o.DTPROJETADO,
        o.REVISAO
    FROM PR_DESENHOROTEIROOPCAO o
    JOIN PR_LISTAENG le ON le.CODIGO = o.LISTAENG
    JOIN PR_DESENHO d ON d.CODIGO = le.CODDESENHO
    WHERE d.DESENHO = '3809029008'
      AND le.POSICAO = '00'
    ORDER BY o.CODIGO DESC
    """

    q_ops = """
    SELECT TOP 20
        op.CODIGO AS OP,
        op.NROOP,
        op.OPSTATUS,
        s.DESCRICAO AS Status_OP,
        op.DESENHO,
        op.DTCADASTRO,
        op.DTINICIO
    FROM PR_OP op
    LEFT JOIN PR_OPSTATUS s ON s.CODIGO = op.OPSTATUS
    WHERE op.DESENHO = '3809029008'
    ORDER BY op.CODIGO DESC
    """

    print("\n=== DESENHO ===")
    print(pd.read_sql(q_desenho, engine).to_string(index=False))

    print("\n=== LISTAENG RAIZ (POSICAO 00) ===")
    print(pd.read_sql(q_lista_raiz, engine).to_string(index=False))

    print("\n=== ROTEIRO OPCAO (RAIZ) ===")
    print(pd.read_sql(q_roteiro_opcao, engine).to_string(index=False))

    print("\n=== OPS DO DESENHO ===")
    print(pd.read_sql(q_ops, engine).to_string(index=False))


if __name__ == "__main__":
    main()
