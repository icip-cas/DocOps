#!/bin/bash
set +e

mkdir -p /logs/verifier

if [ -f /root/original_task_description.txt ]; then
  cp /root/original_task_description.txt /logs/verifier/original_task_description.txt || true
fi

if [ -f /root/task_metadata.json ]; then
  cp /root/task_metadata.json /logs/verifier/task_metadata.json || true
fi

pytest --ctrf /logs/verifier/ctrf.json /tests/test_outputs.py -rA -v
PYTEST_EXIT_CODE=$?

if [ -d /root/submission ]; then
  cp -R /root/submission /logs/verifier/submission || true
fi

if [ $PYTEST_EXIT_CODE -eq 0 ]; then
  echo 1 > /logs/verifier/reward.txt
else
  echo 0 > /logs/verifier/reward.txt
fi

exit 0
