"""
SmartLimo AI - Grille de tarifs officielle (Backstage Limousine Orlando)

À exécuter une fois manuellement (`python -m app.seed_demo_pricing` depuis
backend/) pour peupler zones / zone_zipcodes / zone_cities / rates /
hourly_rates / surcharges / wait_time_rules.

Les valeurs ci-dessous sont transcrites directement du PDF tarifaire fourni
par l'entreprise (Backstage Limousine Orlando - rates.pdf), en l'absence des
fichiers Excel d'origine (TEMPLATE_Zones.xlsx, TEMPLATE_Rates.xlsx, etc.)
restés sur le poste de stage. Ce ne sont donc PAS des valeurs de démo
inventées : ce sont les vrais tarifs, sauf pour les combinaisons absentes du
PDF (ex: Port Canaveral -> Legoland), qui restent non couvertes (comme dans
la grille source elle-même, volontairement a trous - voir
import_pricing.py).

Contrairement à import_pricing.py (qui lit des fichiers Excel), ce script
n'a aucune dépendance à pandas/openpyxl : toutes les valeurs sont codées en
dur ci-dessous.

Les codes de zone et de véhicule utilisés ici ne sont PAS arbitraires : ils
sont recopiés des fonctions réellement utilisées par le chatbot
(app/services/reservation_service.py: resolve_pickup_zone_code,
resolve_dropoff_zone_code, VEHICLE_NAME_TO_RATE_CODE - c'est
estimate_price_by_zone, pas estimate_price, qui est appelée par
app/ai/dialogue_manager.py) et de app/fill_postal_codes.py (codes postaux
par zone). Si ces fonctions changent, ce script doit être mis à jour en
conséquence.

Idempotent (upsert par clé naturelle, comme import_pricing.py) : peut être
relancé sans créer de doublons.
"""

from app.database import SessionLocal
from app.models import (
    Zone, ZoneZipcode, ZoneCity, Rate, HourlyRate, Surcharge, WaitTimeRule,
)

# --- Zones ---------------------------------------------------------------
# Code, description, state, [codes postaux], [villes]
# Les codes postaux reprennent exactement fill_postal_codes.py ; MCO/SFB
# n'y figurent pas (aéroports, pas de zipcode résidentiel associé côté
# métier) mais sont bien nécessaires comme zones de pickup.
ZONES = [
    ("MCO", "Orlando International Airport", "FL", [], ["Orlando"]),
    ("SFB", "Orlando Sanford International Airport", "FL", [], ["Sanford"]),
    ("PORT", "Port Canaveral", "FL",
     ["32920", "32931", "32932"], ["Cape Canaveral", "Cocoa Beach"]),
    ("DISNEY - LAKE B VISTA", "Walt Disney World / Lake Buena Vista", "FL",
     ["32830", "32836", "32821"], ["Lake Buena Vista", "Bay Lake"]),
    ("UNIVERSAL", "Universal Orlando Resort", "FL",
     ["32819"], ["Orlando"]),
    ("KISSIMMEE - SOUTH AREA", "Kissimmee - South Area", "FL",
     ["34741", "34742", "34743", "34744", "34745", "34746", "34747", "34758", "34759"],
     ["Kissimmee", "Celebration"]),
    ("DAVENPORT", "Davenport / Champions Gate", "FL",
     ["33836", "33837", "33896", "33897"], ["Davenport", "Champions Gate"]),
    ("LEGOLAND", "Legoland Florida", "FL", ["33884"], ["Winter Haven"]),
]

# --- Tarifs fixes zone-a-zone (grille officielle, prix tout compris) -------
# (zone_from, zone_to) -> {vehicle_code: prix}. Prix transcrits tels quels du
# PDF "PRICING FROM MCO/SANFORD/PORT CANAVERAL AIRPORT" - pas de peage/
# parking/taxe indiques separement dans ce document, d'ou tolls/parking/tax
# a None dans seed_rates ci-dessous.
RATES = {
    ("MCO", "DISNEY - LAKE B VISTA"): {"SEDAN": 110.0, "SUV": 140.0, "VAN": 155.0, "SPRINTER VAN": 195.0, "LIMOUSINE": 210.0},
    ("MCO", "UNIVERSAL"): {"SEDAN": 110.0, "SUV": 140.0, "VAN": 155.0, "SPRINTER VAN": 195.0, "LIMOUSINE": 210.0},
    ("MCO", "KISSIMMEE - SOUTH AREA"): {"SEDAN": 125.0, "SUV": 165.0, "VAN": 175.0, "SPRINTER VAN": 220.0, "LIMOUSINE": 235.0},
    ("MCO", "DAVENPORT"): {"SEDAN": 140.0, "SUV": 185.0, "VAN": 195.0, "SPRINTER VAN": 240.0, "LIMOUSINE": 255.0},
    ("MCO", "PORT"): {"SEDAN": 185.0, "SUV": 205.0, "VAN": 225.0, "SPRINTER VAN": 290.0, "LIMOUSINE": 295.0},
    ("MCO", "LEGOLAND"): {"SEDAN": 220.0, "SUV": 250.0, "VAN": 275.0, "SPRINTER VAN": 320.0, "LIMOUSINE": 335.0},

    ("SFB", "DISNEY - LAKE B VISTA"): {"SEDAN": 170.0, "SUV": 210.0, "VAN": 225.0, "SPRINTER VAN": 285.0, "LIMOUSINE": 295.0},
    ("SFB", "UNIVERSAL"): {"SEDAN": 170.0, "SUV": 210.0, "VAN": 225.0, "SPRINTER VAN": 285.0, "LIMOUSINE": 295.0},
    ("SFB", "KISSIMMEE - SOUTH AREA"): {"SEDAN": 190.0, "SUV": 220.0, "VAN": 245.0, "SPRINTER VAN": 305.0, "LIMOUSINE": 315.0},
    ("SFB", "DAVENPORT"): {"SEDAN": 215.0, "SUV": 235.0, "VAN": 260.0, "SPRINTER VAN": 345.0, "LIMOUSINE": 340.0},
    ("SFB", "PORT"): {"SEDAN": 225.0, "SUV": 260.0, "VAN": 280.0, "SPRINTER VAN": 345.0, "LIMOUSINE": 355.0},
    ("SFB", "LEGOLAND"): {"SEDAN": 260.0, "SUV": 275.0, "VAN": 305.0, "SPRINTER VAN": 400.0, "LIMOUSINE": 400.0},

    # Pas de ligne "Legoland" depuis Port Canaveral dans le PDF source :
    # combinaison non couverte, comme dans la grille officielle elle-meme.
    ("PORT", "DISNEY - LAKE B VISTA"): {"SEDAN": 195.0, "SUV": 225.0, "VAN": 260.0, "SPRINTER VAN": 315.0, "LIMOUSINE": 325.0},
    ("PORT", "UNIVERSAL"): {"SEDAN": 195.0, "SUV": 225.0, "VAN": 260.0, "SPRINTER VAN": 315.0, "LIMOUSINE": 325.0},
    ("PORT", "KISSIMMEE - SOUTH AREA"): {"SEDAN": 220.0, "SUV": 235.0, "VAN": 280.0, "SPRINTER VAN": 335.0, "LIMOUSINE": 350.0},
    ("PORT", "DAVENPORT"): {"SEDAN": 255.0, "SUV": 245.0, "VAN": 290.0, "SPRINTER VAN": 345.0, "LIMOUSINE": 365.0},
}

# --- Tarifs horaires (grille officielle) -----------------------------------
# Codes vehicule tels que documentes dans models.py (HourlyRate.__doc__) et
# dans le PDF ("HOURLY RATES").
HOURLY_RATES = [
    # code, tarif/h, minimum d'heures, preavis (heures)
    ("SEDAN", 70.0, 3, 6),
    ("PREMIUM SEDAN", 150.0, 3, 24),
    ("SUV", 90.0, 3, 6),
    ("PREMIUM SUV", 125.0, 3, 24),
    ("VAN", 125.0, 3, 24),
    ("SPRINTER EXECUTIVE", 150.0, 3, 48),
    ("LIMOUSINE", 125.0, 3, 48),
    ("PARTY BUS", 225.0, 3, 48),
    ("MINI BUS", 125.0, 3, 48),
]
# Regle additionnelle du PDF, non modelisee dans le schema actuel (pas de
# notion de rayon en km dans HourlyRate) : "Hourly service rate will
# increase by 20% if exceeds 35 mile radius." Ajoutee ci-dessous comme
# supplement informatif (HOURLY_RADIUS_SURCHARGE).

# --- Supplements (grille officielle : "ADDITIONAL CHARGES") ----------------
# Les 4 premiers codes sont ceux lus explicitement par
# reservation_service.apply_surcharges : leurs noms doivent rester exacts.
# Les suivants sont informatifs (non lus par le code actuel) mais
# completent la grille officielle.
SURCHARGES = [
    ("EARLY_LATE", "Early-Late Charge (00:00 - 05:00)", 20.0, None,
     "Supplement forfaitaire applique aux prises en charge entre minuit et 5h."),
    ("HOLIDAY_SURCHARGE", "Holiday Surcharge", None, 20.0,
     "Majoration de 20% appliquee les jours feries."),
    ("AIRPORT_FEE_PORT_CANAVERAL", "Airport fee (arrivals only)", 12.0, None,
     "Frais fixes de 12$ pour toute arrivee prise en charge a l'aeroport (MCO/SFB)."),
    ("SERVICE_FEE", "Port Canaveral fee (each way)", 5.0, None,
     "Frais de 5$ chaque sens pour les trajets depuis/vers Port Canaveral."),
    ("CHILD_SEAT", "Child seat", 0.0, None,
     "Sieges enfants et rehausseurs offerts (complimentary)."),
    ("MEET_GREET", "Meet and Greet Fee", None, None,
     "Inclus (pas de supplement)."),
    ("EXTRA_STOP", "Additional Stop", 20.0, None,
     "20$ si l'arret supplementaire est dans la meme direction."),
    ("GRATUITY", "Gratuity", None, None,
     "Le pourboire n'est pas inclus automatiquement."),
    ("HOURLY_RADIUS_SURCHARGE", "Hourly service - radius surcharge", None, 20.0,
     "Majoration de 20% si le service a l'heure depasse un rayon de 35 miles."),
]

# --- Regles de temps d'attente gratuit (grille officielle) ------------------
WAIT_TIME_RULES = [
    # type de trajet, franchise (min), increment (min), tarif par increment
    ("DOMESTIC", 40, 30, 25.0),
    ("INTERNATIONAL", 60, 30, 30.0),
    ("LOCAL", 15, 1, 1.0),
]


def seed_zones(db) -> dict:
    zones_by_code = {}
    for code, description, state, zipcodes, cities in ZONES:
        zone = db.query(Zone).filter(Zone.code == code).first()
        if zone is None:
            zone = Zone(code=code)
            db.add(zone)
        zone.description = description
        zone.state = state
        db.flush()  # necessaire pour obtenir zone.id avant zipcodes/cities

        # Idempotent par remplacement, comme import_pricing.py.
        db.query(ZoneZipcode).filter(ZoneZipcode.zone_id == zone.id).delete()
        db.query(ZoneCity).filter(ZoneCity.zone_id == zone.id).delete()
        for zipcode in zipcodes:
            db.add(ZoneZipcode(zone_id=zone.id, zipcode=zipcode))
        for city in cities:
            db.add(ZoneCity(zone_id=zone.id, city=city))

        zones_by_code[code] = zone
    db.commit()
    return zones_by_code


def seed_rates(db, zones_by_code: dict) -> int:
    count = 0
    for (from_code, to_code), prices_by_vehicle in RATES.items():
        zone_from = zones_by_code[from_code]
        zone_to = zones_by_code[to_code]

        for vehicle_code, price in prices_by_vehicle.items():
            rate = db.query(Rate).filter(
                Rate.vehicle_code == vehicle_code,
                Rate.zone_from_id == zone_from.id,
                Rate.zone_to_id == zone_to.id,
            ).first()
            if rate is None:
                rate = Rate(
                    vehicle_code=vehicle_code,
                    zone_from_id=zone_from.id,
                    zone_to_id=zone_to.id,
                )
                db.add(rate)

            rate.rate = price
            rate.tolls = None
            rate.parking = None
            rate.tax1 = None
            rate.tax2 = None
            rate.matrix = "BACKSTAGE_OFFICIAL"
            rate.is_default_matrix = True
            count += 1
    db.commit()
    return count


def seed_hourly_rates(db) -> int:
    count = 0
    for vehicle_code, hourly_rate, minimum_hours, notice_hours in HOURLY_RATES:
        row = db.query(HourlyRate).filter(HourlyRate.vehicle_code == vehicle_code).first()
        if row is None:
            row = HourlyRate(vehicle_code=vehicle_code)
            db.add(row)
        row.hourly_rate = hourly_rate
        row.minimum_hours = minimum_hours
        row.notice_hours = notice_hours
        count += 1
    db.commit()
    return count


def seed_surcharges(db) -> int:
    count = 0
    for code, label, amount, percent, description in SURCHARGES:
        row = db.query(Surcharge).filter(Surcharge.code == code).first()
        if row is None:
            row = Surcharge(code=code)
            db.add(row)
        row.label = label
        row.amount = amount
        row.percent = percent
        row.description = description
        count += 1
    db.commit()
    return count


def seed_wait_time_rules(db) -> int:
    count = 0
    for trip_type, grace_period_minutes, increment_minutes, rate_per_increment in WAIT_TIME_RULES:
        row = db.query(WaitTimeRule).filter(WaitTimeRule.trip_type == trip_type).first()
        if row is None:
            row = WaitTimeRule(trip_type=trip_type)
            db.add(row)
        row.grace_period_minutes = grace_period_minutes
        row.increment_minutes = increment_minutes
        row.rate_per_increment = rate_per_increment
        count += 1
    db.commit()
    return count


def run():
    db = SessionLocal()
    try:
        zones_by_code = seed_zones(db)
        rates_count = seed_rates(db, zones_by_code)
        hourly_count = seed_hourly_rates(db)
        surcharges_count = seed_surcharges(db)
        wait_time_count = seed_wait_time_rules(db)
        print("Zones :", len(zones_by_code))
        print("Tarifs zone-a-zone :", rates_count)
        print("Tarifs horaires :", hourly_count)
        print("Supplements :", surcharges_count)
        print("Regles de temps d'attente :", wait_time_count)
    finally:
        db.close()


if __name__ == "__main__":
    run()
