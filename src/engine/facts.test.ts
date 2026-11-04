import { describe, expect, it } from "vitest";
import { findFactPlaceholders, renderFactPlaceholders, replaceFactPlaceholders, scanPlaceholders } from "./facts";
import type { Fact } from "./schema";

describe("fact placeholders", () => {
  const text = "{{fact:vat-rate}} და {{ fact:income-tax-rate }}; ისევ {{fact:vat-rate}}. {{fact:VAT}} {{vat}}";

  it("finds well-formed placeholders once each, and flags malformed ones", () => {
    expect(findFactPlaceholders(text)).toEqual(["vat-rate", "income-tax-rate"]);
    expect(scanPlaceholders(text).malformed).toEqual(["{{fact:VAT}}", "{{vat}}"]);
  });

  it("replaces placeholders with whatever the caller renders", () => {
    expect(replaceFactPlaceholders(text, (id) => `[${id}]`)).toBe(
      "[vat-rate] და [income-tax-rate]; ისევ [vat-rate]. {{fact:VAT}} {{vat}}",
    );
  });

  it("renders facts with their display text, else the value as written", () => {
    const base = { text_ka: "ტექსტი", source: "https://x.ge", checked: null, checkedBy: null } as const;
    const facts: Fact[] = [
      { ...base, id: "vat-rate", value: 0.18, display: "18%" },
      { ...base, id: "iban-length", value: 22 },
    ];
    expect(renderFactPlaceholders("{{fact:vat-rate}}, {{fact:iban-length}}, {{fact:nope}}", facts)).toBe("18%, 22, {{fact:nope}}");
  });
});
