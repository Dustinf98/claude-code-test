import json


def load_rules(path="categories.json"):
    with open(path, "r") as f:
        return json.load(f)


_PAYMENT_KEYWORDS = ["PAYMENT THANK YOU", "PAIEMENT MERCI", "PAYMENT/PAIEMENT"]


def categorize(description, rules):
    upper = description.upper()
    # Detect credit-card payments before keyword rules
    for kw in _PAYMENT_KEYWORDS:
        if kw in upper:
            return "Payment"
    for category, keywords in rules.items():
        for keyword in keywords:
            if keyword.upper() in upper:
                return category
    return "Uncategorized"


def save_rules(rules, path="categories.json"):
    with open(path, "w") as f:
        json.dump(rules, f, indent=2)
