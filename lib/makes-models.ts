export interface CarSuggestion {
  label: string;
  make: string;
  model?: string;
}

const MAKES: { make: string; models: string[] }[] = [
  {
    make: "Toyota",
    models: [
      "Hilux", "Hilux 2.4 GD-6 SRX", "Hilux 2.4 GD-6 SR", "Hilux 2.4 GD-6 Raider",
      "Hilux 2.8 GD-6 Raider", "Hilux 2.8 GD-6 Legend", "Hilux 2.8 GD-6 Legend RS",
      "Hilux 2.8 GD-6 GR-S", "Hilux 4.0 V6 Legend",
      "Fortuner", "Fortuner 2.4 GD-6 GX", "Fortuner 2.4 GD-6 GX AT",
      "Fortuner 2.4 GD-6 4x4 MT", "Fortuner 2.4 GD-6 4x4 AT",
      "Fortuner 2.8 GD-6 4x4 AT", "Fortuner 2.8 GD-6 GR-S", "Fortuner 2.8 GD-6 Legend RS",
      "Corolla", "Corolla 1.8 Prestige CVT", "Corolla 1.8 Hybrid XS CVT",
      "Corolla 2.0 XR CVT", "Corolla Quest 1.8",
      "RAV4", "RAV4 2.0 GX CVT", "RAV4 2.0 GX-R", "RAV4 2.5 Hybrid XS",
      "Land Cruiser", "Land Cruiser 200", "Land Cruiser 300", "Land Cruiser 79",
      "Land Cruiser Prado", "Land Cruiser Prado TX", "Land Cruiser Prado VX",
      "Yaris", "Yaris 1.5 Xi", "Yaris 1.5 Xs", "Yaris Cross 1.5 XS CVT",
      "Camry", "Camry 2.5 XSE CVT", "Camry 2.5 Hybrid XSE",
      "HiAce", "HiAce 2.8 GD Panel Van", "HiAce 2.8 GD Ses'fikile",
      "Starlet", "Starlet 1.4 Xi", "Starlet 1.4 Xs MT",
      "Urban Cruiser 1.5 XS CVT",
    ],
  },
  {
    make: "Volkswagen",
    models: [
      "Polo", "Polo 1.0 TSI Life", "Polo 1.0 TSI Style", "Polo 1.0 TSI Trendline",
      "Polo 1.0 TSI Comfortline", "Polo 1.6 TDI Highline", "Polo GTI 2.0 TSI",
      "Polo Vivo", "Polo Vivo 1.4 Trendline", "Polo Vivo 1.4 Comfortline",
      "Polo Vivo 1.4 Highline", "Polo Vivo 1.6 GT", "Polo Vivo 1.0 MPI",
      "Golf", "Golf 8 1.4 TSI Life", "Golf 8 2.0 TSI R-Line", "Golf GTI 2.0 TSI",
      "Golf R 2.0 TSI 4Motion", "Golf 7 1.4 TSI Comfortline",
      "Tiguan", "Tiguan 1.4 TSI Trendline", "Tiguan 1.4 TSI Comfortline",
      "Tiguan 2.0 TDI Highline", "Tiguan 2.0 TSI 4Motion R-Line",
      "Amarok", "Amarok 2.0 BiTDI Highline", "Amarok 3.0 TDI V6 Extreme",
      "Amarok 3.0 TDI V6 Aventura", "Amarok 2.0 TDI Trendline",
      "T-Roc", "T-Roc 1.4 TSI Design", "T-Roc 2.0 TSI R-Line",
      "T-Cross 1.0 TSI Comfortline", "T-Cross 1.5 TSI R-Line",
      "Touareg 3.0 TDI V6 Executive",
      "Caddy 2.0 TDI Panel Van", "Transporter 2.0 TDI Panel Van",
      "Passat 2.0 TDI Elegance",
    ],
  },
  {
    make: "BMW",
    models: [
      "1 Series", "1 Series 118i",  "1 Series 120d", "1 Series M135i xDrive",
      "2 Series", "2 Series 218i Gran Coupe", "2 Series 220i Coupe", "2 Series M240i xDrive",
      "3 Series", "3 Series 318i", "3 Series 320i", "3 Series 320d",
      "3 Series 320d xDrive", "3 Series 330i", "3 Series M340i xDrive",
      "4 Series", "4 Series 420i Coupe", "4 Series 430i Coupe", "4 Series M440i xDrive",
      "5 Series", "5 Series 520i", "5 Series 520d", "5 Series 530i",
      "5 Series 530d", "5 Series 540i xDrive", "5 Series M550i xDrive",
      "7 Series 740i", "7 Series 750i xDrive",
      "X1", "X1 sDrive18i", "X1 sDrive20d", "X1 xDrive20d", "X1 xDrive25i",
      "X2", "X2 sDrive18i", "X2 sDrive20d", "X2 M35i",
      "X3", "X3 sDrive18d", "X3 sDrive20i", "X3 xDrive20d", "X3 xDrive20i",
      "X3 xDrive30i", "X3 xDrive30d", "X3 M40i", "X3 20d",
      "X4", "X4 xDrive20d", "X4 xDrive30i", "X4 M40i",
      "X5", "X5 xDrive30d", "X5 xDrive40i", "X5 xDrive45e", "X5 M50i", "X5 M50d",
      "X6", "X6 xDrive30d", "X6 xDrive40i", "X6 M50i",
      "X7 xDrive30d", "X7 xDrive40i",
      "M2 Competition", "M3 Competition", "M4 Competition", "M5 Competition",
    ],
  },
  {
    make: "Mercedes-Benz",
    models: [
      "A-Class", "A180", "A200", "A200d", "A250", "AMG A35", "AMG A45 S",
      "B-Class", "B200", "B220d",
      "C-Class", "C180", "C200", "C220d", "C300", "C300d", "AMG C43", "AMG C63",
      "E-Class", "E200", "E220d", "E300", "E350d", "AMG E53", "AMG E63 S",
      "S-Class", "S400d", "S500", "S580",
      "CLA", "CLA200", "CLA220d", "AMG CLA35", "AMG CLA45 S",
      "GLA", "GLA200", "GLA220d", "AMG GLA35", "AMG GLA45 S",
      "GLB", "GLB200", "GLB220d", "AMG GLB35",
      "GLC", "GLC200", "GLC220d", "GLC300", "GLC300d", "AMG GLC43", "AMG GLC63 S",
      "GLE", "GLE300d", "GLE350d", "GLE400d", "AMG GLE53", "AMG GLE63 S",
      "GLS", "GLS400d", "GLS450",
      "G-Class", "G400d", "AMG G63",
    ],
  },
  {
    make: "Ford",
    models: [
      "Ranger", "Ranger 2.0 SiT XL", "Ranger 2.0 SiT XLS", "Ranger 2.0 SiT XLT",
      "Ranger 2.0 BiT Wildtrak", "Ranger 2.0 BiT Stormtrak", "Ranger 3.0 V6 Wildtrak",
      "Ranger Raptor 3.0 V6", "Ranger 2.2 TDCi XL", "Ranger 2.2 TDCi XLS",
      "Everest", "Everest 2.0 SiT XLS 4x2", "Everest 2.0 BiT Sport 4WD",
      "Everest 2.0 BiT Titanium 4WD", "Everest 3.0 V6 Platinum 4WD",
      "Fiesta 1.0 EcoBoost Titanium", "Fiesta 1.6 TDCi Titanium",
      "Focus 1.5 EcoBoost ST-Line", "Focus 2.3 ST",
      "EcoSport 1.0 EcoBoost Titanium", "EcoSport 1.5 TDCi Titanium",
      "Territory 1.5 EcoBoost Trend", "Territory 1.5 EcoBoost Titanium",
      "Mustang 2.3 EcoBoost", "Mustang 5.0 V8 GT",
    ],
  },
  {
    make: "Hyundai",
    models: [
      "i20", "i20 1.2 Motion", "i20 1.4 Fluid", "i20 1.4 Motion", "i20 N 1.6 T-GDI",
      "i30 1.4 T-GDI Elite", "i30 N 2.0 T-GDI",
      "Tucson", "Tucson 2.0 Premium", "Tucson 1.6 T-GDI Premium DCT",
      "Tucson 2.0 Elite", "Tucson 1.6 T-GDI 48V Hybrid",
      "Creta", "Creta 1.5 Premium IVT", "Creta 1.5 Executive IVT", "Creta 1.4 T-GDI Sport DCT",
      "ix35 2.0 Premium", "ix35 2.0 GLS",
      "Santa Fe", "Santa Fe 2.2 CRDi Elite AWD", "Santa Fe 2.2 CRDi Premium AWD",
      "Elantra 1.5 MPI Executive IVT", "Elantra N 2.0 T-GDI",
      "Venue 1.0 T-GDI Motion DCT", "Venue 1.0 T-GDI Fluid DCT",
      "Staria 2.2 CRDi Executive", "H-1 2.5 CRDi Wagon",
      "Grand i10 1.0 Motion", "Grand i10 1.2 Fluid",
    ],
  },
  {
    make: "Kia",
    models: [
      "Sportage", "Sportage 2.0 EX", "Sportage 1.6 T-GDI EX DCT",
      "Sportage 2.0 CRDi EX AWD", "Sportage GT-Line 1.6 T-GDI",
      "Seltos 1.5 EX IVT", "Seltos 1.4 T-GDI EX DCT",
      "Sonet 1.0 T-GDI EX DCT", "Sonet 1.5 EX IVT",
      "Cerato", "Cerato 1.6 EX", "Cerato 2.0 EX",
      "Picanto 1.0 LX", "Picanto 1.2 EX",
      "Stinger 2.0 T-GDI GT-Line", "Stinger 3.3 T-GDI GT",
      "Sorento 2.2 CRDi EX AWD", "Sorento 1.6 T-GDI Hybrid",
      "Carnival 2.2 CRDi SXL", "EV6 GT-Line AWD",
    ],
  },
  {
    make: "Audi",
    models: [
      "A1", "A1 1.0 TFSI Sport S tronic", "A1 30 TFSI Advanced",
      "A3", "A3 1.0 TFSI S tronic", "A3 35 TFSI S line", "A3 40 TFSI S line",
      "A3 35 TDI S line", "S3 2.0 TFSI",
      "A4", "A4 2.0 TDI Design", "A4 2.0 TFSI Design", "A4 40 TFSI Advanced",
      "A4 35 TDI Advanced", "S4 3.0 TFSI",
      "A5", "A5 2.0 TFSI Sport S tronic", "A5 40 TFSI S line", "S5 3.0 TFSI",
      "A6", "A6 40 TDI S tronic", "A6 45 TFSI S tronic", "S6 3.0 TFSI",
      "Q2", "Q2 35 TFSI Advanced S tronic",
      "Q3", "Q3 1.4 TFSI S tronic", "Q3 35 TFSI Advanced", "Q3 35 TDI Advanced",
      "Q5", "Q5 40 TDI quattro S tronic", "Q5 45 TFSI quattro S tronic",
      "Q5 55 TFSI e quattro", "SQ5 3.0 TFSI",
      "Q7", "Q7 45 TDI quattro", "Q7 55 TFSI quattro",
      "Q8", "Q8 55 TFSI quattro", "SQ8 4.0 TFSI",
      "TT 45 TFSI Coupe", "RS3 2.5 TFSI", "RS6 4.0 TFSI Avant",
    ],
  },
  {
    make: "Nissan",
    models: [
      "NP300 Hardbody", "NP300 Hardbody 2.5 TDI 4x4", "NP300 Hardbody 2.5 TDI LE",
      "NP300 Hardbody 2.4 SE",
      "Navara", "Navara 2.5 dCi LE 4x4", "Navara 2.5 dCi SE 4x4", "Navara Pro-4X 2.5 dCi",
      "X-Trail", "X-Trail 2.0 Acenta 4x2", "X-Trail 2.5 Acenta 4x4",
      "X-Trail 1.3 DIG-T Acenta", "X-Trail 1.3 DIG-T Tekna",
      "Qashqai", "Qashqai 1.3 DIG-T Acenta", "Qashqai 1.3 DIG-T Tekna",
      "Micra 1.0 Acenta", "Micra 1.0 Tekna",
      "Magnite 1.0 Turbo Acenta", "Magnite 1.0 Turbo Tekna",
      "Almera 1.5 Acenta", "Almera 1.5 Tekna",
      "Patrol 4.0 V6 LE",
    ],
  },
  {
    make: "Mazda",
    models: [
      "Mazda2", "Mazda2 1.5 Dynamic", "Mazda2 1.5 Active",
      "Mazda3", "Mazda3 1.5 Active", "Mazda3 2.0 Carbon Edition", "Mazda3 2.5 Astina",
      "CX-3", "CX-3 2.0 Active", "CX-3 2.0 Individual",
      "CX-30", "CX-30 2.0 Active", "CX-30 2.0 Carbon Edition", "CX-30 2.5 Astina AWD",
      "CX-5", "CX-5 2.0 Active", "CX-5 2.0 Carbon Edition", "CX-5 2.5 Astina AWD",
      "CX-5 2.2D Active AWD", "CX-5 2.2D Akera AWD",
      "CX-60 3.3D Homura AWD",
      "BT-50 3.0 TDCi Active 4x4", "BT-50 3.0 TDCi Thunder 4x4",
    ],
  },
  {
    make: "Honda",
    models: [
      "Jazz", "Jazz 1.2 Comfort CVT", "Jazz 1.5 Comfort CVT",
      "Ballade", "Ballade 1.5 Comfort CVT", "Ballade 1.5 Elegance CVT",
      "Civic", "Civic 1.5 VTEC Turbo Sport", "Civic 1.5 VTEC Turbo RS",
      "HR-V", "HR-V 1.5 Comfort CVT", "HR-V 1.5 Elegance CVT",
      "CR-V", "CR-V 1.5 Comfort AWD CVT", "CR-V 1.5 Elegance AWD CVT",
      "WR-V 1.5 Comfort CVT", "WR-V 1.5 Elegance CVT",
      "Elevate 1.5 Comfort CVT",
    ],
  },
  {
    make: "Suzuki",
    models: [
      "Swift", "Swift 1.2 GL", "Swift 1.2 GLX CVT", "Swift Sport 1.4 Boosterjet",
      "Jimny", "Jimny 1.5 GLX AllGrip", "Jimny 1.5 GA AllGrip",
      "Vitara", "Vitara 1.4 Boosterjet GL AllGrip", "Vitara 1.6 GLX",
      "Grand Vitara 1.5 GL AllGrip",
      "Baleno", "Baleno 1.5 GL CVT", "Baleno 1.5 GLX CVT",
      "Ertiga 1.5 GL CVT", "Ertiga 1.5 GLX CVT",
      "S-Presso 1.0 GL", "S-Presso 1.0 GLX AMT",
      "Fronx 1.5 GL CVT", "Fronx 1.0 Boosterjet GLX",
      "Brezza 1.5 GL CVT",
    ],
  },
  {
    make: "Renault",
    models: [
      "Kwid", "Kwid 1.0 Dynamique", "Kwid 1.0 Climber",
      "Sandero", "Sandero 66kW Dynamique", "Sandero Stepway 66kW Dynamique",
      "Duster", "Duster 1.5 dCi Dynamique", "Duster 1.6 Dynamique CVT",
      "Clio", "Clio 1.0 Turbo Authentique", "Clio 1.0 Turbo Intens",
      "Captur", "Captur 1.0 Turbo Authentique", "Captur 1.0 Turbo Intens",
      "Koleos 2.0 dCi Intens 4WD",
      "Triber 1.0 Dynamique", "Triber 1.0 Intens AMT",
      "Kiger 1.0 Turbo Zen", "Kiger 1.0 Turbo Intens CVT",
    ],
  },
  {
    make: "Isuzu",
    models: [
      "D-Max", "D-Max 1.9 TD LS-U", "D-Max 1.9 TD L360", "D-Max 3.0 TD LX 4x4",
      "D-Max 3.0 TD LS 4x4", "D-Max 3.0 TD X-Rider",
      "mu-X", "mu-X 1.9 TD LS", "mu-X 3.0 TD LS 4x4", "mu-X 3.0 TD LX 4x4",
      "KB 250D LE",
    ],
  },
  {
    make: "Haval",
    models: [
      "H1 1.5 City", "H1 1.5 Sport",
      "H2 1.5T City", "H2 1.5T Sport",
      "H6", "H6 1.5T Comfort DCT", "H6 1.5T Premium DCT", "H6 2.0T Premium 4WD",
      "Jolion", "Jolion 1.5T Comfort DCT", "Jolion 1.5T Premium DCT",
      "F7 1.5T Comfort DCT", "F7 2.0T Premium 4WD",
      "F7x 2.0T Premium 4WD",
    ],
  },
  {
    make: "Land Rover",
    models: [
      "Range Rover", "Range Rover 3.0 SDV6 SE", "Range Rover 5.0 V8 Autobiography",
      "Range Rover 3.0 D300 SE",
      "Range Rover Sport", "Range Rover Sport 3.0 SDV6 SE", "Range Rover Sport 5.0 V8 SVR",
      "Range Rover Sport 3.0 D300 HSE",
      "Range Rover Evoque", "Range Rover Evoque 2.0D SE", "Range Rover Evoque P250 SE",
      "Range Rover Velar", "Range Rover Velar 2.0D SE R-Dynamic",
      "Discovery", "Discovery 3.0 TD6 SE", "Discovery 3.0 D300 S",
      "Discovery Sport 2.0D SE", "Discovery Sport 2.0D R-Dynamic",
      "Defender", "Defender 90 P300 SE", "Defender 110 P300 SE",
      "Defender 110 3.0D SE",
    ],
  },
  {
    make: "Mitsubishi",
    models: [
      "Pajero", "Pajero 3.2 DiD GLS", "Pajero 3.8 V6 GLS",
      "Pajero Sport", "Pajero Sport 2.4D 4WD", "Pajero Sport 3.0 V6 4WD",
      "Outlander 2.0 GLS", "Outlander 2.4 MIVEC GLS AWD",
      "Eclipse Cross 1.5T GLS", "Eclipse Cross 2.2D 4WD",
      "Triton 2.4D 4WD Club Cab",
      "ASX 2.0 GLS CVT",
    ],
  },
  {
    make: "Jeep",
    models: [
      "Wrangler", "Wrangler 3.6 Rubicon", "Wrangler 3.6 Sahara",
      "Wrangler 2.0T Sahara", "Wrangler 4xe PHEV Rubicon",
      "Grand Cherokee", "Grand Cherokee 3.0 CRD Limited", "Grand Cherokee 5.7 V8 Overland",
      "Grand Cherokee 3.6 Laredo",
      "Cherokee 2.0T Longitude Plus",
      "Compass 1.3T Longitude Plus DCT",
      "Renegade 1.3T Longitude DCT",
    ],
  },
  {
    make: "Subaru",
    models: [
      "Outback 2.5 Premium CVT", "Outback 2.5 Sport CVT",
      "Forester 2.0i-S e-Boxer", "Forester 2.0i-L",
      "XV 2.0i Premium CVT", "XV 1.6i Premium",
      "Impreza 2.0i Premium CVT",
      "WRX 2.4 Sport CVT", "WRX STI 2.5",
    ],
  },
  {
    make: "Porsche",
    models: [
      "911 Carrera", "911 Carrera S", "911 Carrera 4S", "911 Turbo S", "911 GT3",
      "Cayenne", "Cayenne 3.0 TDI", "Cayenne Coupe S", "Cayenne Turbo",
      "Macan", "Macan S", "Macan GTS",
      "Panamera", "Panamera 4S", "Panamera Turbo S",
      "718 Boxster", "718 Cayman", "718 Spyder",
      "Taycan 4S", "Taycan Turbo",
    ],
  },
  {
    make: "Peugeot",
    models: [
      "208 1.2 PureTech Active", "208 1.2 PureTech Allure",
      "2008 1.2 PureTech Active", "2008 1.2 PureTech Allure",
      "308 1.6 THP GTi",
      "3008 1.6 THP GT Line",
      "508 2.0 BlueHDi GT",
      "Landtrek 1.9D 4x4",
    ],
  },
  {
    make: "Opel",
    models: [
      "Corsa", "Corsa 1.2T Enjoy", "Corsa 1.2T Edition",
      "Astra 1.4T Sport", "Astra 1.6T OPC",
      "Crossland 1.2T Enjoy", "Crossland 1.2T Edition",
      "Grandland 1.2T Enjoy", "Grandland 1.6T Edition AWD",
      "Mokka 1.2T Enjoy", "Mokka 1.2T Edition",
    ],
  },
  {
    make: "Mini",
    models: [
      "Cooper 1.5 Classic", "Cooper S 2.0 Sport", "Cooper S Works 2.0",
      "Countryman Cooper D", "Countryman Cooper S", "Countryman JCW 2.0",
      "Clubman Cooper S All4",
      "Convertible Cooper S 2.0",
    ],
  },
  {
    make: "Chery",
    models: [
      "Tiggo 4 1.5T Executive DCT", "Tiggo 4 1.5T Comfort",
      "Tiggo 7 1.5T Executive DCT", "Tiggo 7 Pro 1.6T Executive DCT",
      "Tiggo 8 Pro 1.6T Executive",
      "Arrizo 5 1.5T Executive DCT",
      "Omoda 5 1.6T Executive DCT",
    ],
  },
  {
    make: "GWM",
    models: [
      "P-Series 2.0TD LT 4x4", "P-Series 2.0TD Ultra 4x4",
      "Steed 6 2.0 VGT SX 4x4",
      "Cannon 2.0T Lux 4WD",
    ],
  },
  {
    make: "Chevrolet",
    models: [
      "Spark 1.2 LS", "Spark 1.2 LT",
      "Cruze 1.4T LS", "Cruze 1.6 LS",
      "Captiva 2.4 LS", "Captiva 2.2D LT AWD",
      "Trailblazer 2.8D LT 4x4",
    ],
  },
  {
    make: "BYD",
    models: [
      "Atto 3", "Atto 3 Standard Range", "Atto 3 Extended Range",
      "Seal", "Seal 82.56kWh AWD",
      "Dolphin", "Dolphin 44.9kWh", "Dolphin 60.48kWh",
      "Han EV", "Tang EV",
      "Shark 1.5T PHEV",
    ],
  },
  {
    make: "Skoda",
    models: [
      "Octavia", "Octavia 1.0 TSI Ambition", "Octavia 2.0 TDI Style",
      "Octavia RS 2.0 TSI", "Octavia Scout 2.0 TDI 4x4",
      "Fabia 1.0 MPI Active", "Fabia 1.0 TSI Style",
      "Superb", "Superb 2.0 TDI Style DSG", "Superb 3.6 V6 FSI L&K",
      "Kodiaq", "Kodiaq 2.0 TDI Style DSG", "Kodiaq RS 2.0 BiTDI",
      "Karoq 1.5 TSI Style DSG",
      "Kamiq 1.0 TSI Ambition",
    ],
  },
  {
    make: "Volvo",
    models: [
      "XC40", "XC40 T4 Momentum", "XC40 T5 R-Design AWD", "XC40 Recharge Pure Electric",
      "XC60", "XC60 T5 Momentum", "XC60 T8 R-Design AWD PHEV",
      "XC90", "XC90 T6 Momentum AWD", "XC90 T8 Excellence AWD PHEV",
      "V60 T5 Momentum", "S60 T5 Momentum",
      "C40 Recharge Pure Electric",
    ],
  },
  {
    make: "MG",
    models: [
      "MG3", "MG3 Excite", "MG3 Essence",
      "MG5", "MG5 Excite CVT", "MG5 Essence CVT",
      "ZS", "ZS 1.5 Excite CVT", "ZS EV Excite", "ZS EV Essence",
      "HS", "HS 1.5T Excite DCT", "HS 1.5T Essence DCT",
      "MG4 EV Excite", "MG4 EV Essence",
    ],
  },
  {
    make: "Lexus",
    models: [
      "IS 300h", "IS 350 F Sport",
      "ES 300h Luxury", "ES 350 Luxury",
      "NX 300h Luxury", "NX 350h F Sport",
      "RX 350 Luxury", "RX 500h F Sport",
      "UX 250h Luxury",
      "LX 600 Luxury",
    ],
  },
  {
    make: "Jaguar",
    models: [
      "XE 2.0D Pure", "XE 2.0T R-Dynamic",
      "XF 2.0D Pure", "XF 2.0T R-Dynamic",
      "XJ 3.0 V6 Premium Luxury",
      "F-Pace 2.0D Pure", "F-Pace 3.0D S AWD",
      "E-Pace 2.0D SE", "E-Pace P250 R-Dynamic",
      "F-Type 2.0T Coupe", "F-Type 5.0 V8 SVR",
      "I-Pace EV400 SE",
    ],
  },
  {
    make: "Alfa Romeo",
    models: [
      "Giulia 2.0T Super", "Giulia 2.9 V6 Quadrifoglio",
      "Stelvio 2.0T Super AWD", "Stelvio 2.9 V6 Quadrifoglio AWD",
      "Tonale 1.5T MHEV Sprint",
    ],
  },
  {
    make: "Fiat",
    models: [
      "500 1.2 Lounge", "500 1.4 Abarth",
      "500X 1.4T Cross Plus AWD",
      "Tipo 1.4 Pop", "Tipo 1.6 Lounge",
      "Panda 1.2 Easy",
    ],
  },
  {
    make: "Citroën",
    models: [
      "C3 1.2 PureTech Feel", "C3 1.2 PureTech Shine",
      "C4 1.2 PureTech Feel", "C4 e-C4 Electric",
      "C5 Aircross 1.6T Feel",
    ],
  },
  {
    make: "Seat",
    models: [
      "Ibiza 1.0 TSI Style", "Ibiza 1.6 TDI FR",
      "Leon 1.4 TSI FR", "Leon 2.0 TSI Cupra",
      "Ateca 1.4 TSI Style", "Ateca 2.0 TDI FR 4Drive",
    ],
  },
  {
    make: "Dodge",
    models: [
      "Challenger 5.7 V8 RT", "Challenger 6.4 V8 Scat Pack", "Challenger 6.2 SRT Hellcat",
      "Charger 5.7 V8 RT", "Charger 6.4 V8 Scat Pack",
      "Durango 5.7 V8 RT AWD",
    ],
  },
  {
    make: "RAM",
    models: [
      "1500 5.7 V8 Laramie Crew Cab",
      "1500 TRX 6.2 V8",
    ],
  },
  {
    make: "Mahindra",
    models: [
      "Pik Up 2.2 mHawk S10", "Pik Up 2.2 mHawk S11 4x4",
      "Scorpio 2.2 mHawk S11",
      "XUV300 1.2T W8",
      "Thar 2.2 mHawk 4x4",
    ],
  },
  {
    make: "Ssangyong",
    models: [
      "Rexton 2.2 4WD SX", "Rexton 2.2 4WD Luxury",
      "Musso 2.2 Luxury 4x4",
      "Korando 1.5T Quartz",
      "Tivoli 1.5T Quartz",
    ],
  },
  {
    make: "Omoda",
    models: [
      "5 1.6T Executive DCT",
      "5 EV",
    ],
  },
  {
    make: "Jaecoo",
    models: [
      "7 1.6T Executive AWD",
    ],
  },
  {
    make: "Jetour",
    models: [
      "X70 1.5T Luxury",
      "Dashing 1.5T",
    ],
  },
  {
    make: "GAC",
    models: [
      "GS3 1.5T Comfort", "GS3 1.5T Sport",
      "GS4 1.5T Comfort", "GS4 1.5T Sport",
    ],
  },
];

// Flat list: each make + each make+model combo
export const ALL_SUGGESTIONS: CarSuggestion[] = MAKES.flatMap(({ make, models }) => [
  { label: make, make },
  ...models.map(model => ({ label: `${make} ${model}`, make, model })),
]);

export function getSuggestions(query: string, limit = 8): CarSuggestion[] {
  const q = query.trim().toLowerCase();
  if (!q) return [];
  return ALL_SUGGESTIONS
    .filter(s => s.label.toLowerCase().includes(q))
    .slice(0, limit);
}
