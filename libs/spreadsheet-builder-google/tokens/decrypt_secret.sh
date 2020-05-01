#!/bin/sh

# --batch to prevent interactive command --yes to assume "yes" for questions
gpg --quiet --batch --yes --decrypt --passphrase="$LARGE_SECRET_PASSPHRASE" \
--output libs/spreadsheet-builder-google/tokens/StoredCredential \
libs/spreadsheet-builder-google/tokens/StoredCredential.gpg