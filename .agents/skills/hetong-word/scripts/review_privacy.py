"""Find obvious private-document referrals in newly authored Word review text.

This deliberately reports categories, not the matched sensitive text. It is a
guard for obvious disclosures, not a replacement for semantic review.
"""
import re
import unicodedata


RULES = (
    ('internal_document', re.compile(
        r'(?:内部|本版|本次|另附|配套)(?:[^。；;!?！？]{0,24})'
        r'(?:报告|底稿|核验记录|修订说明|问题清单)|'
        r'(?:审核|审查|复核|风险分析)报告|内部(?:底稿|台账|核验记录)')),
    ('document_referral', re.compile(
        r'(?:详见|参见|参阅|另见|查阅|参考|参照|按照|根据|按|见)'
        r'[^。；;!?！？]{0,80}(?:报告|另一份文件|另一份文档|其他文件|其他文档|修订说明|核验记录)|'
        r'(?:详见|参见|参阅|另见|查阅)[《“\"]|'
        r'(?:see|referto|accordingto)(?:the)?[^。；;!?！？]{0,60}(?:report|reviewmemo|reviewnotes|anotherfile)')),
    ('internal_index', re.compile(
        r'(?:详见|参见|查看|依据|根据|对应|按|见)[《“\"]?(?:r|q|c)-\d{2,}(?!\d)')),
    ('document_filename', re.compile(r'\.(?:docx?|pdf|xlsx?|pptx?|md|html?)(?:$|[^a-z0-9])')),
    ('private_path_or_link', re.compile(
        r'(?:file://|sandbox:|codex://|project-file:|library-file:|/users/|/home/|'
        r'[a-z]:[\\/]|~/|\\\\[a-z0-9]|docs\.google\.com/|notion\.(?:so|site)/|sharepoint\.com/)')),
)


def privacy_findings(text):
    """Return rule codes; normalize fullwidth/spacing without altering input."""
    normalized = unicodedata.normalize('NFKC', text).casefold()
    normalized = re.sub(r'[\s\u200b-\u200d\ufeff]+', '', normalized)
    return [code for code, pattern in RULES if pattern.search(normalized)]


def require_review_privacy(text, label='新增审阅文字'):
    codes = privacy_findings(text)
    if codes:
        raise ValueError(f'{label}违反Word审阅隐私铁律：不得转介另一份文件或写入内部文件名、路径、链接或内部文件索引（稳定问题编号除外）；请就地写清意见（{", ".join(codes)}）')
