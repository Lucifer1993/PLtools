from obfuscator.msdocument import MSDocument


class Modifier:
    def run(self, doc: MSDocument) -> None:
        raise NotImplementedError()


class Pipe:
    def __init__(self, doc: MSDocument):
        self.doc = doc

    def run(self, *args: Modifier) -> None:
        for m in args:
            m.run(self.doc)
