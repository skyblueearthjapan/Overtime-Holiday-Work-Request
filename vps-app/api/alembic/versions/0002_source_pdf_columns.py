"""requests に source_pdf_file_id・source_pdf_folder_id（ステージングへの移行で空にした本番の PDF の ID の控え）

Revision ID: 0002
Revises: 0001
Create Date: 2026-11-26
"""
from __future__ import annotations

from alembic import op
import sqlalchemy as sa

revision = '0002'
down_revision = '0001'
branch_labels = None
depends_on = None


def upgrade() -> None:
    with op.batch_alter_table('requests', schema=None) as batch_op:
        batch_op.add_column(sa.Column('source_pdf_file_id', sa.Text(), server_default='', nullable=False))
        batch_op.add_column(sa.Column('source_pdf_folder_id', sa.Text(), server_default='', nullable=False))


def downgrade() -> None:
    with op.batch_alter_table('requests', schema=None) as batch_op:
        batch_op.drop_column('source_pdf_folder_id')
        batch_op.drop_column('source_pdf_file_id')
