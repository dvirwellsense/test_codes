import test_config as cfg


def run(metadata, **kwargs):

    if cfg.EXPECTED_REF_CAPS is None:
        return True, "RefCaps check disabled"

    ref_caps = metadata.get("RefCaps")

    if ref_caps is None:
        return False, "RefCaps field missing from metadata"

    if ref_caps != cfg.EXPECTED_REF_CAPS:
        return False, (
            f"RefCaps={ref_caps} "
            f"(expected {cfg.EXPECTED_REF_CAPS})"
        )

    return True, f"RefCaps={ref_caps} as expected"
