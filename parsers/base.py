"""
Abstract base class for all bank statement parsers.
"""
from abc import ABC, abstractmethod
from models.statement import Statement


class BaseParser(ABC):
    """
    All parsers must implement :meth:`parse` and return a :class:`Statement`.
    """

    @abstractmethod
    def parse(self, file_path: str) -> Statement:
        """
        Parse *file_path* and return a populated :class:`Statement`.

        :param file_path: Absolute path to the source document (PDF, CSV, …).
        :returns: Populated :class:`~models.statement.Statement` object.
        """
        raise NotImplementedError
