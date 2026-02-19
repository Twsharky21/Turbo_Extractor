Turbo Extractor V3 - Structured Core (Update 1)

This update adds the initial core scaffolding:
- core/errors.py   : AppError + stable error codes
- core/models.py   : Project tree dataclasses + RunReport structs
- core/parsing.py  : Column/row spec parsing + col letter/index utilities

Next step: add core/io.py + a tiny pytest to lock down parsing behavior.
