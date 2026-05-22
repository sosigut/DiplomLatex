from datetime import datetime, timezone

from fastapi import HTTPException, Request, status
from sqlalchemy.orm import Session

from app.core.security import (
    create_access_token,
    create_refresh_token,
    get_refresh_token_expiry,
    get_refresh_token_hash,
    hash_password,
    verify_password,
)
from app.models.department import Department
from app.models.faculty import Faculty
from app.models.refresh_token import RefreshToken
from app.models.user import User

ALLOWED_ROLES = {
    "Доцент",
    "Профессор",
    "Старший преподаватель",
    "Преподаватель",
    "Аспирант",
}


def register_user(db: Session, fio: str, email: str, password: str, faculty_code: str, department_name: str, role: str) -> User:
    existing_user = db.query(User).filter(User.email == email).first()
    if existing_user:
        raise HTTPException(status_code=400, detail="Пользователь с таким email уже существует")

    if role not in ALLOWED_ROLES:
        raise HTTPException(status_code=400, detail="Недопустимая роль")

    faculty = db.query(Faculty).filter(Faculty.faculty_code == faculty_code).first()
    if not faculty:
        raise HTTPException(status_code=404, detail="Факультет с таким кодом не найден")

    department = db.query(Department).filter(Department.department_name == department_name).first()
    if not department:
        raise HTTPException(status_code=404, detail="Кафедра с таким названием не найдена")

    if department.id_faculty != faculty.id_faculty:
        raise HTTPException(status_code=400, detail="Кафедра не принадлежит указанному факультету")

    user = User(
        fio=fio,
        email=email,
        password_hash=hash_password(password),
        role=role,
        id_faculty=faculty.id_faculty,
        id_department=department.id_department,
    )

    db.add(user)
    db.commit()
    db.refresh(user)
    return user


def login_user(db: Session, email: str, password: str, request: Request) -> tuple[str, str]:
    user = db.query(User).filter(User.email == email).first()
    if not user or not verify_password(password, user.password_hash):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Неверный email или пароль"
        )

    access_token = create_access_token(user_id=user.id_user, email=user.email)
    refresh_token = create_refresh_token()

    refresh_record = RefreshToken(
        token_hash=get_refresh_token_hash(refresh_token),
        expires_at=get_refresh_token_expiry(),
        revoked=False,
        user_agent=request.headers.get("user-agent"),
        ip_address=request.client.host if request.client else None,
        id_user=user.id_user,
    )

    db.add(refresh_record)
    db.commit()

    return access_token, refresh_token


def refresh_user_token(db: Session, refresh_token: str) -> str:
    token_hash = get_refresh_token_hash(refresh_token)

    token_record = (
        db.query(RefreshToken)
        .filter(RefreshToken.token_hash == token_hash)
        .first()
    )

    if not token_record:
        raise HTTPException(status_code=401, detail="Refresh token не найден")

    if token_record.revoked:
        raise HTTPException(status_code=401, detail="Refresh token отозван")

    expires_at = token_record.expires_at
    if expires_at.tzinfo is None:
        expires_at = expires_at.replace(tzinfo=timezone.utc)

    if expires_at < datetime.now(timezone.utc):
        raise HTTPException(status_code=401, detail="Refresh token истёк")

    user = db.query(User).filter(User.id_user == token_record.id_user).first()
    if not user:
        raise HTTPException(status_code=401, detail="Пользователь не найден")

    return create_access_token(user_id=user.id_user, email=user.email)


def logout_user(db: Session, refresh_token: str):
    token_hash = get_refresh_token_hash(refresh_token)

    token_record = (
        db.query(RefreshToken)
        .filter(RefreshToken.token_hash == token_hash)
        .first()
    )

    if not token_record:
        raise HTTPException(status_code=404, detail="Refresh token не найден")

    token_record.revoked = True
    db.commit()


def logout_all_user_sessions(db: Session, user_id: int):
    tokens = db.query(RefreshToken).filter(
        RefreshToken.id_user == user_id,
        RefreshToken.revoked == False
    ).all()

    for token in tokens:
        token.revoked = True

    db.commit()