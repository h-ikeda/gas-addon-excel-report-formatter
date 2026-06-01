# Shared pytest configuration for the Cloud Function tests.
#
# The function no longer reads a bundled template from disk — the template is
# supplied in the request body — so no working-directory fixture is required.
