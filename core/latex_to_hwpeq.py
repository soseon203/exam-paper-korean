"""LaTeX 수식을 HWP 수식 스크립트로 변환하는 모듈.

한글(HWP)의 수식 편집기는 자체 수식 스크립트 문법을 사용합니다.
이 모듈은 LaTeX 수식을 HWP 수식 스크립트로 변환합니다.

주요 매핑:
  \\frac{a}{b}     → {a} over {b}
  x^{2}            → x ^{2}
  x_{n}            → x _{n}
  \\sqrt{x}         → sqrt {x}
  \\sqrt[n]{x}      → root {n} of {x}
  \\sum_{i=0}^{n}   → SUM _{i=0} ^{n}
  \\int_{a}^{b}     → INT _{a} ^{b}
  \\lim_{x \\to 0}   → lim _{x -> 0}
"""

from __future__ import annotations

import re


class LaTeXToHWPConverter:
    """LaTeX → HWP 수식 스크립트 변환기."""

    # 그리스 문자 매핑
    GREEK_MAP = {
        r"\alpha": "alpha",
        r"\beta": "beta",
        r"\gamma": "gamma",
        r"\delta": "delta",
        r"\epsilon": "epsilon",
        r"\varepsilon": "varepsilon",
        r"\zeta": "zeta",
        r"\eta": "eta",
        r"\theta": "theta",
        r"\vartheta": "vartheta",
        r"\iota": "iota",
        r"\kappa": "kappa",
        r"\lambda": "lambda",
        r"\mu": "mu",
        r"\nu": "nu",
        r"\xi": "xi",
        r"\pi": "pi",
        r"\rho": "rho",
        r"\sigma": "sigma",
        r"\tau": "tau",
        r"\upsilon": "upsilon",
        r"\phi": "phi",
        r"\varphi": "varphi",
        r"\chi": "chi",
        r"\psi": "psi",
        r"\omega": "omega",
        # 대문자 (PascalCase — HWP 규칙: 첫 글자만 대문자)
        r"\Gamma": "Gamma",
        r"\Delta": "Delta",
        r"\Theta": "Theta",
        r"\Lambda": "Lambda",
        r"\Xi": "Xi",
        r"\Pi": "Pi",
        r"\Sigma": "Sigma",
        r"\Upsilon": "Upsilon",
        r"\Phi": "Phi",
        r"\Chi": "Chi",
        r"\Psi": "Psi",
        r"\Omega": "Omega",
    }

    # 연산자/기호 매핑
    SYMBOL_MAP = {
        # 산술 연산
        r"\times": "TIMES",
        r"\cdot": "CDOT",
        r"\div": "DIV",
        r"\pm": "PLUSMINUS",
        r"\mp": "MINUSPLUS",
        # 관계 연산
        r"\leq": "LEQ",
        r"\le": "LEQ",
        r"\geq": "GEQ",
        r"\ge": "GEQ",
        r"\neq": "neq",
        r"\ne": "neq",
        r"\approx": "APPROX",
        r"\equiv": "EQUIV",
        r"\sim": "SIM",
        r"\simeq": "SIMEQ",
        r"\cong": "CONG",
        r"\propto": "PROPTO",
        r"\asymp": "ASYMP",
        r"\doteq": "DOTEQ",
        r"\prec": "PREC",
        r"\succ": "SUCC",
        r"\ll": "<<",
        r"\gg": ">>",
        # 특수 기호
        r"\infty": "inf",
        r"\partial": "partial",
        r"\nabla": "LAPLACE",
        r"\forall": "forall",
        r"\exists": "EXIST",
        r"\in": "in",
        r"\notin": "notin",
        r"\ni": "OWNS",
        r"\subset": "subset",
        r"\supset": "supset",
        r"\subseteq": "subseteq",
        r"\supseteq": "supseteq",
        r"\cup": "SMALLUNION",
        r"\cap": "SMALLINTER",
        r"\emptyset": "emptyset",
        r"\vee": "VEE",
        r"\lor": "VEE",
        r"\wedge": "WEDGE",
        r"\land": "WEDGE",
        r"\neg": "LNOT",
        r"\lnot": "LNOT",
        r"\oplus": "OPLUS",
        r"\otimes": "OTIMES",
        r"\therefore": "therefore",
        r"\because": "because",
        r"\angle": "angle",
        r"\perp": "BOT",
        r"\parallel": "parallel",
        r"\triangle": "TRIANGLE",
        r"\square": '"□"',
        r"\circ": "CIRC",
        r"\bullet": "BULLET",
        r"\star": "STAR",
        r"\diamond": "DIAMOND",
        r"\top": "TOP",
        r"\vdash": "VDASH",
        r"\models": "MODELS",
        # 화살표 — 단일선
        r"\rightarrow": "->",
        r"\leftarrow": "<-",
        r"\leftrightarrow": "<->",
        r"\to": "->",
        r"\gets": "<-",
        r"\uparrow": "uparrow",
        r"\downarrow": "downarrow",
        r"\updownarrow": "udarrow",
        r"\nearrow": "nearrow",
        r"\nwarrow": "nwarrow",
        r"\searrow": "searrow",
        r"\swarrow": "swarrow",
        r"\hookleftarrow": "hookleft",
        r"\hookrightarrow": "hookright",
        r"\mapsto": "mapsto",
        # 화살표 — 이중선
        r"\Rightarrow": "RARROW",
        r"\Leftarrow": "LARROW",
        r"\Leftrightarrow": "LRARROW",
        r"\Uparrow": "UPARROW",
        r"\Downarrow": "DOWNARROW",
        r"\Updownarrow": "UDARROW",
        # 점
        r"\ldots": "LDOTS",
        r"\cdots": "CDOTS",
        r"\vdots": "VDOTS",
        r"\ddots": "DDOTS",
        # 기타 기호
        r"\prime": "prime",
        r"\aleph": "ALEPH",
        r"\hbar": "HBAR",
        r"\imath": "IMATH",
        r"\jmath": "JMATH",
        r"\ell": "ELL",
        r"\wp": "WP",
        r"\Im": "IMAG",
        r"\Re": "REIMAGE",
        r"\dagger": "DAGGER",
        r"\ddagger": "DDAGGER",
    }

    # 함수 매핑
    FUNC_MAP = {
        r"\sin": "sin",
        r"\cos": "cos",
        r"\tan": "tan",
        r"\sec": "sec",
        r"\csc": "csc",
        r"\cot": "cot",
        r"\cosec": "cosec",
        r"\arcsin": "arcsin",
        r"\arccos": "arccos",
        r"\arctan": "arctan",
        r"\sinh": "sinh",
        r"\cosh": "cosh",
        r"\tanh": "tanh",
        r"\coth": "coth",
        r"\log": "log",
        r"\ln": "ln",
        r"\lg": "lg",
        r"\exp": "exp",
        r"\Exp": "Exp",
        r"\det": "det",
        r"\max": "max",
        r"\min": "min",
        r"\sup": "sup",
        r"\inf": "inf",
        r"\lim": "lim",
        r"\Lim": "Lim",
        r"\gcd": "gcd",
        r"\arg": "arg",
        r"\dim": "dim",
        r"\ker": "ker",
        r"\hom": "hom",
        r"\mod": "mod",
        r"\lcm": "lcm",
    }

    # 장식(accent) 매핑
    ACCENT_MAP = {
        r"\vec": "VEC",
        r"\bar": "BAR",
        r"\hat": "HAT",
        r"\tilde": "TILDE",
        r"\dot": "DOT",
        r"\ddot": "DDOT",
        r"\acute": "acute",
        r"\grave": "grave",
        r"\check": "check",
        r"\breve": "arch",
        r"\overline": "overline",
        r"\underline": "underline",
        r"\overrightarrow": "VEC",
        r"\widehat": "HAT",
        r"\widetilde": "TILDE",
    }

    def __init__(self):
        self._build_patterns()

    def _build_patterns(self):
        """정규식 패턴 사전 컴파일."""
        # \frac{a}{b}
        self._frac_pattern = re.compile(
            r"\\frac\s*" + self._brace_group("num") + r"\s*" + self._brace_group("den")
        )
        # \sqrt[n]{x} 또는 \sqrt{x}
        self._sqrt_n_pattern = re.compile(
            r"\\sqrt\s*\[([^\]]+)\]\s*" + self._brace_group("body")
        )
        self._sqrt_pattern = re.compile(r"\\sqrt\s*" + self._brace_group("body"))

        # \sum, \prod, \int 등 대형 연산자
        self._big_op_pattern = re.compile(
            r"\\(sum|prod|coprod|int|iint|iiint|oint|bigcup|bigcap)"
            r"(?:\s*_\s*" + self._brace_group_or_char("lo") + r")?"
            r"(?:\s*\^\s*" + self._brace_group_or_char("hi") + r")?"
        )

        # accent: \vec{A}, \bar{x} 등
        accent_cmds = "|".join(
            re.escape(k[1:]) for k in self.ACCENT_MAP
        )
        self._accent_pattern = re.compile(
            r"\\(" + accent_cmds + r")\s*" + self._brace_group("body")
        )

        # \left( ... \right)
        self._leftright_pattern = re.compile(
            r"\\left\s*([(\[{|.])\s*(.*?)\s*\\right\s*([)\]}|.])",
            re.DOTALL,
        )

        # 상첨자/하첨자
        self._superscript = re.compile(r"\^\s*" + self._brace_group_or_char("sup"))
        self._subscript = re.compile(r"_\s*" + self._brace_group_or_char("sub"))

        # \text{...}
        self._text_pattern = re.compile(r"\\text\s*" + self._brace_group("txt"))
        # \mathrm{...}
        self._mathrm_pattern = re.compile(r"\\mathrm\s*" + self._brace_group("txt"))
        # \mathbf{...}
        self._mathbf_pattern = re.compile(r"\\mathbf\s*" + self._brace_group("txt"))

        # \binom{n}{k}
        self._binom_pattern = re.compile(
            r"\\binom\s*" + self._brace_group("top") + r"\s*" + self._brace_group("bot")
        )

        # \begin{env}...\end{env} (행렬/조건식)
        self._env_pattern = re.compile(
            r"\\begin\{(cases|pmatrix|bmatrix|vmatrix|matrix)\}"
            r"\s*(.*?)\s*"
            r"\\end\{\1\}",
            re.DOTALL,
        )

    @staticmethod
    def _brace_group(name: str) -> str:
        """Named group for {content} - handles up to 3 levels of nesting."""
        # 각 단계가 한 단계 더 깊은 중괄호를 허용
        L0 = r"[^{}]*"
        L1 = r"(?:[^{}]|\{" + L0 + r"\})*"
        L2 = r"(?:[^{}]|\{" + L1 + r"\})*"
        L3 = r"(?:[^{}]|\{" + L2 + r"\})*"
        return r"\{(?P<" + name + r">" + L3 + r")\}"

    @staticmethod
    def _brace_group_or_char(name: str) -> str:
        """Named group for {content} or single char."""
        L0 = r"[^{}]*"
        L1 = r"(?:[^{}]|\{" + L0 + r"\})*"
        L2 = r"(?:[^{}]|\{" + L1 + r"\})*"
        L3 = r"(?:[^{}]|\{" + L2 + r"\})*"
        return (
            r"(?:\{(?P<" + name + r">" + L3 + r")\}"
            r"|(?P<" + name + r"_c>[^\s{}\\]))"
        )

    def _get_match(self, match: re.Match, name: str) -> str:
        """brace_group_or_char에서 값 추출."""
        val = match.group(name)
        if val is None:
            val = match.group(name + "_c")
        return val or ""

    def convert(self, latex: str) -> str:
        """LaTeX 수식을 HWP 수식 스크립트로 변환.

        Args:
            latex: LaTeX 수식 문자열

        Returns:
            HWP 수식 스크립트 문자열

        Raises:
            ValueError: 변환 실패 시
        """
        # 전처리: 불필요한 공백, $기호 제거
        s = latex.strip().strip("$").strip()

        # displaymath 환경 제거
        for env in [r"\[", r"\]", r"\(", r"\)"]:
            s = s.replace(env, "")
        for env_name in ["equation", "align", "gather", "displaymath"]:
            s = re.sub(r"\\begin\{" + env_name + r"\*?\}", "", s)
            s = re.sub(r"\\end\{" + env_name + r"\*?\}", "", s)

        s = s.strip()
        result = self._convert_expr(s)

        # 후처리: 다중 공백 정리
        result = re.sub(r"  +", " ", result).strip()
        return result

    def _convert_expr(self, s: str) -> str:
        """재귀적으로 LaTeX 표현식을 변환."""
        if not s:
            return ""

        # 0. 행렬/조건식 환경: \begin{env}...\end{env}
        def _env_repl(m: re.Match) -> str:
            env = m.group(1)
            content = m.group(2)
            env_map = {
                "cases": "CASES",
                "pmatrix": "PMATRIX",
                "bmatrix": "BMATRIX",
                "vmatrix": "DMATRIX",
                "matrix": "MATRIX",
            }
            hwp_env = env_map[env]
            # \\ → # (행 구분자 변환)
            content = re.sub(r"\\\\", " # ", content)
            content = self._convert_expr(content)
            return hwp_env + " {" + content + "}"

        s = self._env_pattern.sub(_env_repl, s)

        # 1. \text, \mathrm, \mathbf
        s = self._text_pattern.sub(lambda m: '"' + m.group("txt") + '"', s)
        s = self._mathrm_pattern.sub(lambda m: "rm " + m.group("txt"), s)
        s = self._mathbf_pattern.sub(lambda m: "bold " + m.group("txt"), s)

        # 2. \binom{n}{k}
        s = self._binom_pattern.sub(
            lambda m: "LEFT ( {"
            + self._convert_expr(m.group("top"))
            + "} atop {"
            + self._convert_expr(m.group("bot"))
            + "} RIGHT )",
            s,
        )

        # 3. \frac{a}{b}
        s = self._frac_pattern.sub(
            lambda m: "{"
            + self._convert_expr(m.group("num"))
            + "} over {"
            + self._convert_expr(m.group("den"))
            + "}",
            s,
        )

        # 4. \sqrt[n]{x} 또는 \sqrt{x}
        s = self._sqrt_n_pattern.sub(
            lambda m: "root {"
            + self._convert_expr(m.group(1))
            + "} of {"
            + self._convert_expr(m.group("body"))
            + "}",
            s,
        )
        s = self._sqrt_pattern.sub(
            lambda m: "sqrt {" + self._convert_expr(m.group("body")) + "}", s
        )

        # 5. 대형 연산자
        def _big_op_repl(m: re.Match) -> str:
            op = m.group(1).upper()
            op_map = {
                "SUM": "SUM", "PROD": "PROD", "COPROD": "COPROD",
                "INT": "INT", "IINT": "DINT", "IIINT": "TINT", "OINT": "OINT",
                "BIGCUP": "UNION", "BIGCAP": "INTER",
            }
            hwp_op = op_map.get(op, op)
            lo = self._get_match(m, "lo")
            hi = self._get_match(m, "hi")
            result = hwp_op
            if lo:
                result += " _{" + self._convert_expr(lo) + "}"
            if hi:
                result += " ^{" + self._convert_expr(hi) + "}"
            return result

        s = self._big_op_pattern.sub(_big_op_repl, s)

        # 6. \left( ... \right)
        def _leftright_repl(m: re.Match) -> str:
            left = m.group(1)
            body = m.group(2)
            right = m.group(3)
            # 구분 문자 매핑
            delim_map = {
                "(": "(", ")": ")", "[": "[", "]": "]",
                r"\{": "lbrace", r"\}": "rbrace",
                "{": "lbrace", "}": "rbrace",
                "|": "|", ".": "",
            }
            l_str = delim_map.get(left, left)
            r_str = delim_map.get(right, right)
            inner = self._convert_expr(body)
            if l_str and r_str:
                return f"LEFT {l_str} {inner} RIGHT {r_str}"
            elif l_str:
                return f"LEFT {l_str} {inner}"
            elif r_str:
                return f"{inner} RIGHT {r_str}"
            return inner

        s = self._leftright_pattern.sub(_leftright_repl, s)

        # 7. accent: \vec{A} → VEC A
        def _accent_repl(m: re.Match) -> str:
            cmd = "\\" + m.group(1)
            body = m.group("body")
            hwp_accent = self.ACCENT_MAP.get(cmd, m.group(1).upper())
            return hwp_accent + " {" + self._convert_expr(body) + "}"

        s = self._accent_pattern.sub(_accent_repl, s)

        # 8. 그리스 문자
        for latex_cmd, hwp_name in sorted(
            self.GREEK_MAP.items(), key=lambda x: -len(x[0])
        ):
            s = s.replace(latex_cmd, hwp_name)

        # 9. 기호/연산자
        for latex_cmd, hwp_sym in sorted(
            self.SYMBOL_MAP.items(), key=lambda x: -len(x[0])
        ):
            s = s.replace(latex_cmd, hwp_sym)

        # 10. 함수명
        for latex_cmd, hwp_func in sorted(
            self.FUNC_MAP.items(), key=lambda x: -len(x[0])
        ):
            s = s.replace(latex_cmd, hwp_func)

        # 11. 상첨자/하첨자 (braces 유지)
        s = self._superscript.sub(
            lambda m: " ^{" + self._convert_expr(self._get_match(m, "sup")) + "}", s
        )
        s = self._subscript.sub(
            lambda m: " _{" + self._convert_expr(self._get_match(m, "sub")) + "}", s
        )

        # 12. { } 내부 재귀 처리 (단순 그룹)
        def _brace_recurse(m: re.Match) -> str:
            inner = m.group(1)
            return "{" + self._convert_expr(inner) + "}"

        s = re.sub(r"\{([^{}]+)\}", _brace_recurse, s)

        # 13. HWP 공백 문자 및 기타 남은 LaTeX 명령어 정리
        s = s.replace("\\,", "`")
        s = s.replace("\\;", "~")
        s = s.replace("\\!", "")
        s = s.replace("\\qquad", "~~~~")
        s = s.replace("\\quad", "~~")
        s = s.replace("\\\\", "")
        s = re.sub(r"\\[a-zA-Z]+", "", s)  # 남은 알 수 없는 명령어 제거

        return s


# 모듈 레벨 싱글톤
_converter = LaTeXToHWPConverter()


def latex_to_hwpeq(latex: str) -> str:
    """LaTeX 수식을 HWP 수식 스크립트로 변환.

    Args:
        latex: LaTeX 수식 문자열 (예: r"\\frac{1}{2}")

    Returns:
        HWP 수식 스크립트 (예: "{1} over {2}")
    """
    return _converter.convert(latex)


def latex_to_image(latex: str, dpi: int = 150) -> bytes:
    """LaTeX 수식을 PNG 이미지로 렌더링 (폴백용).

    Args:
        latex: LaTeX 수식 문자열
        dpi: 이미지 DPI

    Returns:
        PNG 이미지 바이트
    """
    import io
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(0.01, 0.01))
    ax.axis("off")

    # matplotlib의 LaTeX 렌더링
    text = ax.text(
        0.5, 0.5,
        f"${latex}$",
        transform=ax.transAxes,
        fontsize=14,
        ha="center", va="center",
    )

    # 텍스트 크기에 맞게 그림 크기 조정
    fig.canvas.draw()
    renderer = fig.canvas.get_renderer()
    bbox = text.get_window_extent(renderer=renderer)
    # 포인트 → 인치 변환 + 여백
    width = bbox.width / dpi + 0.1
    height = bbox.height / dpi + 0.1
    fig.set_size_inches(max(width, 0.5), max(height, 0.3))

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", pad_inches=0.02,
                transparent=True)
    plt.close(fig)
    buf.seek(0)
    return buf.read()
