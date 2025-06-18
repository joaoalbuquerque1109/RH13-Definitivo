# test_app.py
import os
import pytest
from db_logic import create_table, add_person, list_people, delete_person, DB_NAME

@pytest.fixture(scope="module", autouse=True)
def setup_and_teardown():
    # Antes dos testes
    if os.path.exists(DB_NAME):
        os.remove(DB_NAME)
    create_table()
    yield
    # Depois dos testes
    if os.path.exists(DB_NAME):
        os.remove(DB_NAME)

def test_add_person():
    person_data = (
        "João Victor", "22", "01-01-2002", "1234567", "12345678901", "Rua X", "Centro", "Goiana",
        "PE", "Fulano", "81999999999", "joao@email.com", "Solteiro", "1234567890", "1234567890",
        "100", "200", "Técnico", "TI", "01-06-2020", "3000", "CLT", "Diurno"
    )
    add_person(person_data)
    results = list_people()
    assert len(results) == 1
    assert results[0][1] == "João Victor"
    assert results[0][5] == "12345678901"  # CPF

def test_delete_person():
    people = list_people()
    person_id = people[0][0]
    delete_person(person_id)
    results = list_people()
    assert len(results) == 0
