VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaDataObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public country As Variant
Public region_fias_id As Variant
Public region_kladr_id As Variant
Public region_with_type As Variant
Public region_type As Variant
Public region_type_full As Variant
Public region As Variant
Public area_fias_id As Variant
Public area_kladr_id As Variant
Public area_with_type As Variant
Public area_type As Variant
Public area_type_full As Variant
Public area As Variant
Public city_fias_id As Variant
Public city_kladr_id As Variant
Public city_with_type As Variant
Public city_type As Variant
Public city_type_full As Variant
Public city As Variant
Public city_area As Variant
Public city_district As Variant
Public settlement_fias_id As Variant
Public settlement_kladr_id As Variant
Public settlement_with_type As Variant
Public settlement_type As Variant
Public settlement_type_full As Variant
Public settlement As Variant
Public street_fias_id As Variant
Public street_kladr_id As Variant
Public street_with_type As Variant
Public street_type As Variant
Public street_type_full As Variant
Public street As Variant
Public house_fias_id As Variant
Public house_kladr_id As Variant
Public house_type As Variant
Public house_type_full As Variant
Public house As Variant
Public block_type As Variant
Public block_type_full As Variant
Public block As Variant
Public flat_type As Variant
Public flat_type_full As Variant
Public flat As Variant
Public flat_area As Variant
Public square_meter_price As Variant
Public flat_price As Variant
Public postal_box As Variant
Public fias_id As Variant
Public fias_level As Variant
Public kladr_id As Variant
Public capital_marker As Variant
Public okato As Variant
Public oktmo As Variant
Public tax_office As Variant
Public tax_office_legal As Variant
Public timezone As Variant
Public geo_lat As Variant
Public geo_lon As Variant
Public beltway_hit As Variant
Public beltway_distance As Variant
Public qc_geo As Variant
Public qc_complete As Variant
Public qc_house As Variant
Public qc As Variant
Public unparsed_parts As Variant
Public value As Variant
Public unrestricted_value As Variant

Public Sub InitiateProperties(json As Object)
    value = json("suggestions")("value")
    unrestricted_value = json("suggestions")("unrestricted_value")
    country = json("suggestions")("data")("country")
    region_fias_id = json("suggestions")("data")("region_fias_id")
    region_kladr_id = json("suggestions")("data")("region_kladr_id")
    region_with_type = json("suggestions")("data")("region_with_type")
    region_type = json("suggestions")("data")("region_type")
    region_type_full = json("suggestions")("data")("region_type_full")
    region = json("suggestions")("data")("region")
    area_fias_id = json("suggestions")("data")("area_fias_id")
    area_kladr_id = json("suggestions")("data")("area_kladr_id")
    area_with_type = json("suggestions")("data")("area_with_type")
    area_type = json("suggestions")("data")("area_type")
    area_type_full = json("suggestions")("data")("area_type_full")
    area = json("suggestions")("data")("area")
    city_fias_id = json("suggestions")("data")("city_fias_id")
    city_kladr_id = json("suggestions")("data")("city_kladr_id")
    city_with_type = json("suggestions")("data")("city_with_type")
    city_type = json("suggestions")("data")("city_type")
    city_type_full = json("suggestions")("data")("city_type_full")
    city = json("suggestions")("data")("city")
    city_area = json("suggestions")("data")("city_area")
    city_district = json("suggestions")("data")("city_district")
    settlement_fias_id = json("suggestions")("data")("settlement_fias_id")
    settlement_kladr_id = json("suggestions")("data")("settlement_kladr_id")
    settlement_with_type = json("suggestions")("data")("settlement_with_type")
    settlement_type = json("suggestions")("data")("settlement_type")
    settlement_type_full = json("suggestions")("data")("settlement_type_full")
    settlement = json("suggestions")("data")("settlement")
    street_fias_id = json("suggestions")("data")("street_fias_id")
    street_kladr_id = json("suggestions")("data")("street_kladr_id")
    street_with_type = json("suggestions")("data")("street_with_type")
    street_type = json("suggestions")("data")("street_type")
    street_type_full = json("suggestions")("data")("street_type_full")
    street = json("suggestions")("data")("street")
    house_fias_id = json("suggestions")("data")("house_fias_id")
    house_kladr_id = json("suggestions")("data")("house_kladr_id")
    house_type = json("suggestions")("data")("house_type")
    house_type_full = json("suggestions")("data")("house_type_full")
    house = json("suggestions")("data")("house")
    block_type = json("suggestions")("data")("block_type")
    block_type_full = json("suggestions")("data")("block_type_full")
    block = json("suggestions")("data")("block")
    flat_type = json("suggestions")("data")("flat_type")
    flat_type_full = json("suggestions")("data")("flat_type_full")
    flat = json("suggestions")("data")("flat")
    flat_area = json("suggestions")("data")("flat_area")
    square_meter_price = json("suggestions")("data")("square_meter_price")
    flat_price = json("suggestions")("data")("flat_price")
    postal_box = json("suggestions")("data")("postal_box")
    fias_id = json("suggestions")("data")("fias_id")
    fias_level = json("suggestions")("data")("fias_level")
    kladr_id = json("suggestions")("data")("kladr_id")
    capital_marker = json("suggestions")("data")("capital_marker")
    okato = json("suggestions")("data")("okato")
    oktmo = json("suggestions")("data")("oktmo")
    tax_office = json("suggestions")("data")("tax_office")
    tax_office_legal = json("suggestions")("data")("tax_office_legal")
    timezone = json("suggestions")("data")("timezone")
    geo_lat = json("suggestions")("data")("geo_lat")
    geo_lon = json("suggestions")("data")("geo_lon")
    beltway_hit = json("suggestions")("data")("beltway_hit")
    beltway_distance = json("suggestions")("data")("beltway_distance")
    qc_geo = json("suggestions")("data")("qc_geo")
    qc_complete = json("suggestions")("data")("qc_complete")
    qc_house = json("suggestions")("data")("qc_house")
    qc = json("suggestions")("data")("qc")
    unparsed_parts = json("suggestions")("data")("unparsed_parts")
End Sub
