#!/bin/sh
# Run each test of test/test_office.vim in a fresh Vim and Neovim (or the editors given as arguments).
set -u
cd "$(dirname "$0")/.." || exit 1
root=$(pwd)
result=$(mktemp)
trap 'rm -f "$result"' EXIT
[ $# -gt 0 ] || set -- vim nvim
failed=0
timeout=
command -v timeout >/dev/null && timeout='timeout -k 10 120'
tests=$(sed -n 's/^function! \(Test_[A-Za-z0-9_]*\)().*/\1/p' test/test_office.vim)
for editor; do
  command -v "$editor" >/dev/null || { echo "SKIP $editor: not installed"; continue; }
  for t in $tests; do
    : > "$result"
    # shellcheck disable=SC2086
    OFFICE_TEST_RESULT=$result $timeout "$editor" -Nu NONE -i NONE -n -es \
      --cmd "set rtp^=$root" -c "let g:test = '$t'" -S test/test_office.vim < /dev/null > /dev/null 2>&1
    line=$(cat "$result")
    [ -n "$line" ] || line='FAIL: no result'
    case $line in FAIL*) failed=$((failed + 1)) ;; esac
    echo "$editor $t: $line"
  done
done
[ "$failed" -eq 0 ]
