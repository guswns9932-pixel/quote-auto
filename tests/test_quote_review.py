"""견적작성 대상 검토 — 금액(파일 생성과 같은 식), 도식 칸의 상태, 점검 목록."""
import quote_calc
import quote_config as qc
import quote_review as rv


def _it(cat, spec, qty, price, role="ITEM", amt=None):
    return {"cat": cat, "role": role, "spec": spec, "qty": qty, "price": price,
            "amt": qty * price if amt is None else amt}


CLASS = {"GV100": "Gate Valve", "MB100": "Metal Bellows", "HN3": "Hot N2 Transfer",
         "RACK3": "Pump Rack", "PDB": "Power Distribution Box", "SB100": "성형 Bellows",
         "OPT": "Pump Options"}


def test_amounts_follow_the_template_formulas():
    a = quote_calc.domestic([_it("PUMP", "P", 3, 13900000), _it("RACK", "R", 3, 264000)], 3)
    assert (a.g15, a.g16, a.unit, a.total, a.unit_check) == (41700000, 792000, 14164000, 42492000, 14164000)
    # Pump Credit 은 PUMP 단가에서, Rack Credit 은 RACK 합계에서 빠진다
    b = quote_calc.domestic([_it("PUMP", "P", 3, 13900000), _it("RACK", "R", 3, 264000),
                             _it("RACK", "Pump Credit", 0, 0, "CREDIT", -1000000),
                             _it("RACK", "Rack Credit", 0, 0, "CREDIT", -500000)], 3)
    assert (b.p_unit_net, b.pump_credit_unit, b.g16) == (13566667, -333333, 292000)
    assert (b.unit, b.total, b.unit_check) == (13664000, 40992000, 13665000)
    assert quote_calc.domestic([], 0).qty == 1 and quote_calc.domestic([], "x").qty == 1


def test_diagram_slots_and_checks():
    items = [_it("PUMP", "P1200", 3, 100), _it("RACK", "RACK3", 1, 50), _it("RACK", "GV100", 2, 10),
             _it("RACK", "HN3", 3, 5), _it("RACK", "OPT", 3, 0), _it("RACK", "X-new", 1, 7)]
    r = rv.analyze(items, "국내", 3, class_of=CLASS, known_specs=set(CLASS) | {"P1200"},
                   dup_price_specs={"HN3"}, suggested=[("MB100", 3), ("RACK3", 1)])
    flow = {s.label: s for s in r.flow}
    assert [s.label for s in r.flow] == ["Gate Valve", "Metal Bellows", "PUMP", "Hot N2 Transfer"]
    assert (flow["Gate Valve"].state, flow["Gate Valve"].note) == (rv.DIFF, "수량 2 ≠ 3CH")
    assert flow["Metal Bellows"].state == rv.MISS and flow["Metal Bellows"].suggest == ("MB100", 3)
    assert (flow["PUMP"].state, flow["PUMP"].specs) == (rv.OK, ["P1200"])
    assert [s.state for s in r.base] == [rv.OK, rv.BLANK, rv.BLANK]      # Pump Rack 은 CH 대조 안 함
    assert [(s.text, s.state) for s in r.others] == [("OPT", rv.BAD), ("X-new", rv.OK)]

    texts = [(c.level, c.text) for c in r.checks]
    assert (rv.RED, "단가 0원: OPT") in texts
    assert (rv.YELLOW, "Gate Valve 수량 2 ≠ 3CH") in texts
    assert (rv.YELLOW, "품목 시트에 단가가 둘 이상인 규격: HN3") in texts
    assert (rv.INFO, "코드매핑에 있는데 안 담김: MB100 × 3") in texts
    assert not any("RACK3" in t for _, t in texts)                        # 이미 담긴 제안은 빠진다
    assert (r.used, r.capacity) == (5, 25) and (r.amounts.g15, r.amounts.g16) == (300, 92)


def test_unknown_spec_capacity_and_empty_list():
    assert rv.analyze([], "국내", 3).checks == []                          # 빈 목록은 지적하지 않는다
    many = [_it("RACK", "R%d" % i, 1, 1) for i in range(26)]
    r = rv.analyze(many + [_it("RACK", "없는것", 1, 0)], "국내", None, known_specs=set())
    texts = [c.text for c in r.red]
    assert "PUMP 가 담기지 않았습니다" in texts
    assert any("칸 25행 초과 (27행)" in t for t in texts)
    assert "단가 0원: 없는것 — 품목 시트에 없는 이름" in texts
    assert r.ch is None and not r.yellow                                  # 수량을 모르면 CH 대조를 안 한다


def test_capacity_by_quote_type_and_settings():
    items = [_it("PUMP", "P%d" % i, 1, 1) for i in range(2)] + [_it("RACK", "R%d" % i, 1, 1) for i in range(21)]
    us = rv.analyze(items, "미국", 1)
    assert any("미국 견적서 PUMP 칸 1행 초과" in c.text for c in us.red)
    assert any("미국 견적서 RACK 칸 20행 초과" in c.text for c in us.red) and us.amounts is None
    cn = rv.analyze(items[:1] + items[2:], "중국", 1)
    assert [c.text for c in cn.red if "초과" in c.text] == ["중국 견적서 RACK 칸 20행 초과 (21행) — 이대로는 생성되지 않습니다"]
    # 도식 구성과 칸 수는 설정을 따른다
    cfg = qc.QuoteConfig({"diagram.flow": "@PUMP*, Gate Valve", "diagram.base": "Pump Rack", "dom.spec_end": 20, "dom.sign_end": 14})
    r = rv.analyze([_it("PUMP", "P", 2, 1), _it("RACK", "GV100", 1, 1)], "국내", 2, cfg=cfg, class_of=CLASS)
    assert [s.label for s in r.flow] == ["PUMP", "Gate Valve"] and r.flow[1].state == rv.OK and r.capacity == 4
