# Version 1.0.0 local validation

40 synthetic engine tests passed on Windows with Python 3.12. Tests cover arbitrary cohort sizes, entity/customer IDs, separate file columns, Excel worksheets, literal ID handling, explicit normalization, collisions, header/encoding choices, exclusions, sample matching, cancellation and failure without completed output.

The bundled demo has five selected IDs and nine expected output records. GUI and frozen-executable tests use fictional data only. The Windows runtime and relocated ZIP are checked without external Python. Actual macOS/Linux and institutional workstation execution remain unverified.

This is local functional verification of table extraction. No real clinical, customer or other private datasets are included or evaluated.
