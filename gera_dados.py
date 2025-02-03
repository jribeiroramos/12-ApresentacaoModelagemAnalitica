import sqlite3
import random
import datetime

def create_db():
    """
    Conecta (ou cria) o banco SQLite `cobertura_daily.db`.
    """
    conn = sqlite3.connect("cobertura_daily.db")
    return conn

def drop_tables(conn):
    """
    Remove as tabelas se existirem, para recriar do zero.
    """
    c = conn.cursor()
    c.execute("DROP TABLE IF EXISTS Fato_Cobertura")
    c.execute("DROP TABLE IF EXISTS Dim_Tempo")
    c.execute("DROP TABLE IF EXISTS Dim_Municipio")
    c.execute("DROP TABLE IF EXISTS Dim_Estrategia")
    conn.commit()

def create_tables(conn):
    """
    Cria as tabelas do Star Schema simplificado (diário).
    """
    c = conn.cursor()

    # Tabela de dimensão Tempo (1 dia = 1 registro)
    c.execute("""
    CREATE TABLE Dim_Tempo (
        ID_Tempo INTEGER PRIMARY KEY,
        DataCompleta TEXT,  -- 'YYYY-MM-DD'
        Ano INTEGER,
        Mes INTEGER,
        Dia INTEGER
    )
    """)

    # Tabela de dimensão Município
    c.execute("""
    CREATE TABLE Dim_Municipio (
        ID_Munic INTEGER PRIMARY KEY,
        Nome_Municipio TEXT,
        Estado TEXT,
        Regiao TEXT
    )
    """)

    # Tabela de dimensão Estratégia
    c.execute("""
    CREATE TABLE Dim_Estrategia (
        ID_Estrategia INTEGER PRIMARY KEY,
        Nome_Estrategia TEXT,
        Tipo_Estrategia TEXT,
        Equipe_Responsavel TEXT
    )
    """)

    # Tabela de Fato (cobertura diária)
    c.execute("""
    CREATE TABLE Fato_Cobertura (
        ID_Fato INTEGER PRIMARY KEY AUTOINCREMENT,
        ID_Tempo INTEGER,
        ID_Munic INTEGER,
        ID_Estrategia INTEGER,
        Coberto INTEGER,  -- 0 ou 1
        Data_Inicio_Cobertura TEXT,
        Data_Fim_Cobertura TEXT,
        FOREIGN KEY (ID_Tempo) REFERENCES Dim_Tempo(ID_Tempo),
        FOREIGN KEY (ID_Munic) REFERENCES Dim_Municipio(ID_Munic),
        FOREIGN KEY (ID_Estrategia) REFERENCES Dim_Estrategia(ID_Estrategia)
    )
    """)

    conn.commit()

def generate_daily_dates(start_date, end_date):
    """
    Gera todas as datas diárias entre start_date (inclusive) e end_date (inclusive).
    Retorna uma lista de datetime.date.
    """
    dates = []
    delta = datetime.timedelta(days=1)
    current = start_date
    while current <= end_date:
        dates.append(current)
        current += delta
    return dates

def populate_dim_tempo(conn):
    """
    Gera datas diárias de 2005-01-01 até 2024-12-31 (20 anos ~ 7305 dias).
    Insere na Dim_Tempo.
    """
    c = conn.cursor()

    start_date = datetime.date(2005, 1, 1)
    end_date = datetime.date(2024, 12, 31)

    all_days = generate_daily_dates(start_date, end_date)

    # ID_Tempo inicia em 1
    id_tempo = 1
    for d in all_days:
        data_str = d.strftime("%Y-%m-%d")
        c.execute("""
            INSERT INTO Dim_Tempo (ID_Tempo, DataCompleta, Ano, Mes, Dia)
            VALUES (?, ?, ?, ?, ?)
        """, (id_tempo, data_str, d.year, d.month, d.day))
        id_tempo += 1

    conn.commit()

def populate_dim_municipio(conn):
    """
    Gera 100 municípios fictícios.
    """
    c = conn.cursor()

    lista_estados = [("SP","Sudeste"),("RJ","Sudeste"),("MG","Sudeste"),("ES","Sudeste"),
                     ("RS","Sul"),("SC","Sul"),("PR","Sul"),
                     ("BA","Nordeste"),("PE","Nordeste"),("CE","Nordeste"),
                     ("AM","Norte"),("PA","Norte"),("MT","Centro-Oeste")]

    for i in range(1, 101):
        nome_municipio = f"Municipio_{i}"
        estado, regiao = random.choice(lista_estados)
        c.execute("""
            INSERT INTO Dim_Municipio (ID_Munic, Nome_Municipio, Estado, Regiao)
            VALUES (?, ?, ?, ?)
        """, (i, nome_municipio, estado, regiao))

    conn.commit()

def populate_dim_estrategia(conn):
    """
    Cria 5 estratégias fictícias.
    """
    c = conn.cursor()
    estrategias = [
        (1, "Expansao Norte", "Marketing Direto", "Equipe Norte"),
        (2, "Parceria Local", "Alianças", "Equipe Parcerias"),
        (3, "Cobertura Digital", "Marketing Online", "Equipe Digital"),
        (4, "Franquia Regional", "Modelo Franquia", "Equipe Franquia"),
        (5, "Expansao Sul", "Vendas Diretas", "Equipe Sul")
    ]
    for estr in estrategias:
        c.execute("""
            INSERT INTO Dim_Estrategia (ID_Estrategia, Nome_Estrategia, Tipo_Estrategia, Equipe_Responsavel)
            VALUES (?, ?, ?, ?)
        """, estr)

    conn.commit()

def get_id_tempo_map(conn):
    """
    Retorna um dicionário que mapeia date (YYYY-MM-DD) -> ID_Tempo,
    e também um array ordenado com (ID_Tempo, date).
    """
    c = conn.cursor()
    c.execute("SELECT ID_Tempo, DataCompleta FROM Dim_Tempo ORDER BY ID_Tempo")
    rows = c.fetchall()
    date_to_id = {}
    id_date_list = []
    for (id_t, data_str) in rows:
        date_obj = datetime.datetime.strptime(data_str, "%Y-%m-%d").date()
        date_to_id[date_obj] = id_t
        id_date_list.append((id_t, date_obj))
    return date_to_id, id_date_list

def random_date_in_range(start_date, end_date):
    """
    Retorna uma data aleatória entre start_date e end_date (inclusive).
    """
    delta = end_date - start_date
    days_range = delta.days
    rand_days = random.randint(0, days_range)
    return start_date + datetime.timedelta(days=rand_days)

def populate_fato_cobertura(conn):
    """
    Para cada município, sorteia:
      - Data_Inicio (entre 2005-01-01 e 2024-12-31)
      - Com 20% de chance, Data_Fim (após início) 
      - Estratégia (1 de 5)
    Para cada dia do período, insere na Fato_Cobertura com Coberto=1 se day está em [inicio, fim),
    senão 0. 
    Armazena Data_Inicio_Cobertura e Data_Fim_Cobertura no registro diário (mesmo que repetitivo).
    """
    c = conn.cursor()

    # Obter mapeamentos de data <-> ID_Tempo
    date_to_id, id_date_list = get_id_tempo_map(conn)

    # Definir range total de datas
    global_start = datetime.date(2005, 1, 1)
    global_end   = datetime.date(2024, 12, 31)

    # Obter municípios
    c.execute("SELECT ID_Munic FROM Dim_Municipio ORDER BY ID_Munic")
    municipios = [row[0] for row in c.fetchall()]

    # Obter estratégias
    c.execute("SELECT ID_Estrategia FROM Dim_Estrategia ORDER BY ID_Estrategia")
    estrategias = [row[0] for row in c.fetchall()]

    inserts = []
    total_records = 0

    for munic in municipios:
        # 1. Sorteia data_inicio entre 2005-01-01 e 2024-12-31
        data_inicio = random_date_in_range(global_start, global_end)
        # 2. Decide se terá data_fim (20% de chance)
        has_fim = (random.random() < 0.20)  # 20% prob
        data_fim = None
        if has_fim:
            # sorteia data_fim entre data_inicio e global_end
            if data_inicio < global_end:
                data_fim = random_date_in_range(data_inicio, global_end)
                # garantir que data_fim > data_inicio 
                # (se for igual, cobertura termina no mesmo dia que inicia)
                if data_fim == data_inicio:
                    # incrementa 1 dia
                    data_fim = data_inicio + datetime.timedelta(days=1)
                # se, por acaso, data_fim sai fora do range, ajusta:
                if data_fim > global_end:
                    data_fim = global_end
            else:
                data_fim = data_inicio  # sem efeito (cobertura praticamente não existirá se já é fim)

        # 3. Sorteia estratégia
        estr_escolhida = random.choice(estrategias)

        # Para cada dia da Dim_Tempo, define se coberto=1 ou 0
        for (id_t, date_obj) in id_date_list:
            if date_obj < data_inicio:
                coberto = 0
            else:
                if data_fim is None:
                    # se não tem data_fim -> permanente
                    coberto = 1
                else:
                    # coberto se date_obj < data_fim
                    if date_obj < data_fim:
                        coberto = 1
                    else:
                        coberto = 0
            
            # Montar strings para armazenar
            data_inicio_str = data_inicio.strftime("%Y-%m-%d")
            data_fim_str = data_fim.strftime("%Y-%m-%d") if data_fim else None

            inserts.append((id_t, munic, estr_escolhida, coberto, data_inicio_str, data_fim_str))
            total_records += 1

    # Inserir no banco
    # Pode ser demorado, dependendo do tamanho (cerca de 730k linhas).
    c.executemany("""
        INSERT INTO Fato_Cobertura 
        (ID_Tempo, ID_Munic, ID_Estrategia, Coberto, Data_Inicio_Cobertura, Data_Fim_Cobertura)
        VALUES (?, ?, ?, ?, ?, ?)
    """, inserts)
    conn.commit()
    print(f"[populate_fato_cobertura] Inseridos {total_records} registros na Fato_Cobertura.")

def main():
    conn = create_db()
    drop_tables(conn)
    create_tables(conn)

    print("Populando Dim_Tempo (diário, 20 anos)...")
    populate_dim_tempo(conn)

    print("Populando Dim_Municipio (100 municípios)...")
    populate_dim_municipio(conn)

    print("Populando Dim_Estrategia (5 estratégias)...")
    populate_dim_estrategia(conn)

    print("Populando Fato_Cobertura (diário, com data_fim opcional)...")
    populate_fato_cobertura(conn)

    conn.close()
    print("Concluído. Dados gerados em 'cobertura_daily.db'.")

if __name__ == "__main__":
    main()

