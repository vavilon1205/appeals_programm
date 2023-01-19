class Case:
    """Перечисление падежей"""
    # Родительный
    GENITIVE = 0
    # Дательный
    DATIVE = 1
    # Винительный
    ACCUSATIVE = 2
    # Творительный
    INSTRUMENTAL = 3
    # Предложный
    PREPOSITIONAL = 4

    CASES = (
        DATIVE,
        GENITIVE,
        ACCUSATIVE,
        INSTRUMENTAL,
        PREPOSITIONAL,
    )


class Gender:
    """Перечисление родов"""
    # Мужской род
    MALE = 'male'
    # Женский род
    FEMALE = 'female'
    # Средний род
    ANDRGN = 'androgynous'
