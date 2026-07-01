import test_config as cfg


def run(metadata, **kwargs):

    missing = [
        key for key in cfg.REQUIRED_METADATA_KEYS
        if key not in metadata
    ]

    if missing:
        return False, f"Missing metadata fields: {missing}"

    return True, f"All {len(cfg.REQUIRED_METADATA_KEYS)} required fields present"
