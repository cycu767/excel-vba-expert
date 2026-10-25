"""Run tests offline; --enforce-os installs a Linux seccomp network-deny filter.

The enforced mode needs Linux libseccomp.so.2, not a Python package. This is a
verification-only dependency; ordinary lookup/import needs only Python stdlib.
The filter applies to this process and descendants and lasts until they exit.
"""
from __future__ import annotations
import argparse
import ctypes
import errno
import io
import json
import os
import platform
import socket
import stat
import subprocess
import sys
import time
import unittest
from contextlib import ExitStack
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]


def deny_network_os() -> dict:
    if sys.platform != 'linux':
        raise RuntimeError('--enforce-os requires Linux/libseccomp; not silently replaced by mocks')
    inherited = []
    for fdpath in Path('/proc/self/fd').iterdir():
        try:
            if stat.S_ISSOCK(os.fstat(int(fdpath.name)).st_mode):
                inherited.append(int(fdpath.name))
        except (OSError, ValueError):
            continue
    if inherited:
        raise RuntimeError('Refusing an offline claim with inherited socket descriptors: ' + repr(inherited))
    lib = ctypes.CDLL('libseccomp.so.2', use_errno=True)
    lib.seccomp_init.argtypes = [ctypes.c_uint32]
    lib.seccomp_init.restype = ctypes.c_void_p
    lib.seccomp_syscall_resolve_name.argtypes = [ctypes.c_char_p]
    lib.seccomp_syscall_resolve_name.restype = ctypes.c_int
    lib.seccomp_rule_add.argtypes = [ctypes.c_void_p, ctypes.c_uint32, ctypes.c_int, ctypes.c_uint]
    lib.seccomp_rule_add.restype = ctypes.c_int
    lib.seccomp_load.argtypes = [ctypes.c_void_p]
    lib.seccomp_load.restype = ctypes.c_int
    lib.seccomp_release.argtypes = [ctypes.c_void_p]
    context = lib.seccomp_init(0x7FFF0000)  # allow ordinary local filesystem/process work
    if not context:
        raise RuntimeError('seccomp_init failed')
    blocked = []
    try:
        for name in ('socket','socketpair','connect','accept','accept4','bind','listen',
                     'sendto','sendmsg','sendmmsg','recvfrom','recvmsg','recvmmsg','shutdown',
                     'io_uring_setup'):
            number = lib.seccomp_syscall_resolve_name(name.encode('ascii'))
            if number >= 0:
                status = lib.seccomp_rule_add(context, 0x00050000 | errno.EPERM, number, 0)
                if status != 0:
                    raise RuntimeError('seccomp_rule_add failed: ' + name)
                blocked.append(name)
        status = lib.seccomp_load(context)
        if status != 0:
            raise RuntimeError('seccomp_load failed: ' + str(status))
    finally:
        lib.seccomp_release(context)
    probes = []
    for label, family in (('IPv4', socket.AF_INET), ('IPv6', socket.AF_INET6)):
        try:
            with socket.socket(family, socket.SOCK_STREAM):
                raise RuntimeError('Network socket unexpectedly succeeded')
        except OSError as exc:
            if exc.errno != errno.EPERM:
                raise
            probes.append({'family': label, 'result': 'DENIED_BY_KERNEL', 'errno': exc.errno})
    child = subprocess.run([sys.executable, '-c',
        'import socket,errno\ntry:\n socket.socket()\nexcept OSError as e:\n'
        ' print(e.errno); raise SystemExit(0 if e.errno==errno.EPERM else 3)\n'
        'raise SystemExit(4)'], capture_output=True, text=True, check=True)
    return {'mode': 'LINUX_SECCOMP_KERNEL_ENFORCED', 'blocked_syscalls': blocked,
            'inherited_sockets': inherited, 'socket_probes': probes,
            'child_process_network_denied': child.stdout.strip() == str(errno.EPERM),
            'scope': 'Test process and its descendants, not the entire host or AI service'}


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--enforce-os', action='store_true')
    p.add_argument('--pattern', default='test_*.py')
    p.add_argument('--out', type=Path, help='New JSON report path; never overwrite')
    a = p.parse_args()
    if a.out and a.out.exists():
        p.error('Report exists; choose a new path')
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    start = time.perf_counter()
    try:
        with ExitStack() as stack:
            if a.enforce_os:
                guard = deny_network_os()
            else:
                def denied(*_, **__):
                    raise RuntimeError('Network forbidden by Python test guard')
                for attr in ('socket','create_connection','getaddrinfo'):
                    stack.enter_context(patch.object(socket, attr, side_effect=denied))
                guard = {'mode': 'PYTHON_GUARD_ONLY', 'scope': 'Not an OS-level network isolation claim'}
            suite = unittest.defaultTestLoader.discover(str(ROOT / 'tests'), pattern=a.pattern)
            stream = io.StringIO()
            result = unittest.TextTestRunner(stream=stream, verbosity=2).run(suite)
            print(stream.getvalue())
            report = {'created_at_utc': datetime.now(timezone.utc).isoformat(),
                      'platform': platform.platform(), 'python': platform.python_version(),
                      'guard': guard, 'tests_run': result.testsRun,
                      'passed': result.testsRun-len(result.failures)-len(result.errors)-len(result.skipped),
                      'failures': len(result.failures), 'errors': len(result.errors), 'skipped': len(result.skipped),
                      'status': 'PASS' if result.wasSuccessful() and not result.skipped else 'FAIL_OR_SKIPPED',
                      'elapsed_seconds': round(time.perf_counter()-start, 3),
                      'log': stream.getvalue()}
            if a.out:
                a.out.parent.mkdir(parents=True, exist_ok=True)
                with a.out.open('x', encoding='utf-8') as f:
                    json.dump(report, f, ensure_ascii=False, indent=2); f.write('\n')
            print(json.dumps({k:v for k,v in report.items() if k != 'log'}, indent=2))
            return 0 if report['status'] == 'PASS' else 1
    except (OSError, RuntimeError, subprocess.SubprocessError) as exc:
        print('ERROR: ' + str(exc), file=sys.stderr)
        return 2


if __name__ == '__main__':
    raise SystemExit(main())
