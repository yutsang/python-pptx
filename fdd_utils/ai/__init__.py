"""Where the model meets the workbook: the configuration and the model dropdown, the client that calls a provider, the evidence
contract each account is read into (evidence_contract.py) and the verifier that holds every sentence to it (claims.py, validator.py).

Submodules are imported where they are used (`from fdd_utils.ai import claims`); only configuration, the client and what the deck builder
takes from the old verifier are re-exported here.
"""
from __future__ import annotations

from .config import (  # noqa: F401
    FDDConfig,
    default_model_choice_key,
    get_default_config_path,
    is_provider_ready,
    list_model_choices,
    load_yaml_config,
)
from .client import AIClient  # noqa: F401
from .validator import strip_thinking  # noqa: F401
