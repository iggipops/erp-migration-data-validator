import logging
import logging.handlers
import sys
from pathlib import Path


def setup_logger(log_file_path: str | None = None) -> logging.Logger:
    """
    Set up the application logger (FS 8.11).

    log_file_path=None — used before configuration has loaded (FS 8.1),
    when no log file path is known yet. Attaches a console handler
    (INFO and above) plus an in-memory buffer that captures everything
    at DEBUG and above but never writes to a file. If configuration
    subsequently fails to load, that buffer is simply discarded — the
    console is the only record (FS 8.11.2).

    log_file_path=<path> — used once configuration has loaded. Attaches
    a console handler (INFO+) and a file handler (DEBUG+, full detail)
    at that path. If an in-memory buffer exists from an earlier
    no-argument call, its contents are replayed into the new file
    handler first, so the log file ends up complete from process start
    with no gap (FS 8.11.2), before the buffer is discarded.
    """
    logger = logging.getLogger("erp_migration_data_validator")
    logger.setLevel(logging.DEBUG)

    # Recover any pre-configuration buffer before we clear handlers.
    buffer_handler = next(
        (h for h in logger.handlers if isinstance(h, logging.handlers.MemoryHandler)),
        None,
    )
    logger.handlers.clear()

    formatter = logging.Formatter(
        fmt="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # Console handler — INFO and above. Always present.
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    if log_file_path is None:
        # No log file path known yet — buffer everything instead of
        # writing to a file. capacity/flushLevel are set so this never
        # auto-flushes; only an explicit setup_logger(log_file_path)
        # call later flushes it, and only into a real file handler.
        new_buffer = logging.handlers.MemoryHandler(
            capacity=100_000,
            flushLevel=logging.CRITICAL + 1,
            target=None,
        )
        new_buffer.setLevel(logging.DEBUG)
        new_buffer.setFormatter(formatter)
        logger.addHandler(new_buffer)
        return logger

    # Log file path known — attach it and replay any buffered
    # pre-configuration detail into it first.
    log_path = Path(log_file_path)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    file_handler = logging.FileHandler(log_file_path, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    if buffer_handler is not None:
        buffer_handler.setTarget(file_handler)
        buffer_handler.flush()
        buffer_handler.close()

    logger.addHandler(file_handler)
    return logger


def get_logger() -> logging.Logger:
    """Return the application logger (must call setup_logger first)."""
    return logging.getLogger("erp_migration_data_validator")
