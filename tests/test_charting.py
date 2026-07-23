from winsif_mon.ui.charting import nice_range, padded_range


def test_nice_range_rounds_to_readable_steps():
    minimum, maximum, step = nice_range(7428.79, 12688.94, tick_count=6)

    assert minimum == 6000
    assert maximum == 14000
    assert step == 2000


def test_padded_range_adds_margin_to_flat_values():
    minimum, maximum = padded_range([12.0, 12.0], padding_ratio=0.05)

    assert minimum < 12.0
    assert maximum > 12.0
