import pytest
from login.models import User

@pytest.mark.django_db
def test_user_creation():
    user = User.objects.create_user(
        username = 'pruebaBanco',
        email = 'andres7villa24@gmail.com',
        password = '1234qpl5678#*',
        first_name = 'Diego'

    )
    assert user.username  == "pruebaBanco"


@pytest.mark.django_db
def test_staff_user_creation():
    user = User.objects.create_user(
        username = 'pruebaBanco',
        email = 'andres7villa24@gmail.com',
        password = '1234qpl5678#*',
        first_name = 'Diego',
        is_staff = True

    )
    assert user.is_staff 