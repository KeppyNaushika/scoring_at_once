"""利用者に見せるエラー."""


class UserError(Exception):
    """処理を続けられない理由. 画面側は title と message をそのままダイアログに出す."""

    def __init__(self, title: str, message: str) -> None:
        super().__init__(message)
        self.title = title
        self.message = message
