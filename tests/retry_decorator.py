"""COM Test Retry Decorator — handles transient Excel COM failures.

Excel COM automation is inherently unstable due to:
  - Excel process startup/shutdown race conditions
  - COM marshalling timeouts
  - Workbook open/close race conditions
  - Windows clipboard contention

This decorator retries failed tests automatically.

Usage:
  from tests.retry_decorator import com_retry

  @com_retry(max_retries=2, delay=1.0)
  def test_stats_mean():
      ...

  # Or as context manager:
  from tests.retry_decorator import retry_com_call
  result = retry_com_call(lambda: xl_app.Run("StatsUtils_Mean", data))
"""

import time
import functools
import logging

logger = logging.getLogger(__name__)

# Common transient COM errors
TRANSIENT_ERRORS = (
    -2147418111,  # RPC_E_CALL_REJECTED
    -2146777998,  # RPC_E_SYS_CALL_FAILED
    -2147023174,  # RPC_E_SERVER_UNAVAILABLE
    -2147024894,  # STG_E_FILENOTFOUND (workbook race)
    -2146827864,  # Excel internal
)


class COMRetryError(Exception):
    """Raised when all retries are exhausted."""
    def __init__(self, func_name, attempts, last_error):
        self.func_name = func_name
        self.attempts = attempts
        self.last_error = last_error
        super().__init__(
            f"{func_name}: failed after {attempts} attempts. "
            f"Last error: {last_error}"
        )


def com_retry(max_retries=2, delay=1.0, backoff=2.0, exceptions=None):
    """Decorator: retry a function on transient COM errors.

    Args:
        max_retries: Number of retry attempts (total calls = max_retries + 1)
        delay: Initial delay between retries in seconds
        backoff: Multiplier for delay after each retry
        exceptions: Tuple of exception types to catch (default: Exception)
    """
    if exceptions is None:
        exceptions = (Exception,)

    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            current_delay = delay
            last_error = None

            for attempt in range(max_retries + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    last_error = e
                    # Check if it's a transient COM error
                    error_code = getattr(e, 'hresult', None)
                    is_transient = (
                        error_code in TRANSIENT_ERRORS
                        or 'COM' in str(type(e).__name__)
                        or 'RPC' in str(e)
                        or 'Call was rejected' in str(e)
                    )

                    if not is_transient and attempt > 0:
                        # Non-transient error on retry — give up
                        raise

                    if attempt < max_retries:
                        logger.warning(
                            f"[RETRY] {func.__name__} attempt {attempt+1}/{max_retries+1} "
                            f"failed: {e}. Retrying in {current_delay:.1f}s..."
                        )
                        time.sleep(current_delay)
                        current_delay *= backoff
                    else:
                        logger.error(
                            f"[FAIL] {func.__name__} exhausted {max_retries+1} attempts."
                        )

            raise COMRetryError(func.__name__, max_retries + 1, last_error)

        return wrapper
    return decorator


def retry_com_call(callable_fn, max_retries=2, delay=1.0):
    """Retry a single COM call (non-decorator form).

    Args:
        callable_fn: Zero-argument callable to execute
        max_retries: Number of retries
        delay: Delay between retries

    Returns:
        Result of callable_fn()

    Raises:
        COMRetryError: If all retries exhausted
    """
    @com_retry(max_retries=max_retries, delay=delay)
    def _inner():
        return callable_fn()

    return _inner()


# ── Test statistics ──────────────────────────────────────────────────────

_retry_stats = {"total_retries": 0, "total_failures": 0}


def get_retry_stats():
    """Get retry statistics for the test session."""
    return dict(_retry_stats)


def reset_retry_stats():
    """Reset retry statistics."""
    _retry_stats["total_retries"] = 0
    _retry_stats["total_failures"] = 0
