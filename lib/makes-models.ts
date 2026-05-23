export interface CarSuggestion {
  label: string;
  make: string;
  model?: string;
}

const MAKES: { make: string; models: string[] }[] = [
  { make: "Toyota", models: ["Hilux", "Fortuner", "Corolla", "Corolla Quest", "RAV4", "Land Cruiser", "Land Cruiser Prado", "Yaris", "Camry", "HiAce", "Quantum", "Starlet", "Urban Cruiser"] },
  { make: "Volkswagen", models: ["Polo", "Polo Vivo", "Golf", "Tiguan", "Amarok", "T-Roc", "T-Cross", "Touareg", "Caddy", "Transporter", "Passat", "Arteon"] },
  { make: "BMW", models: ["1 Series", "2 Series", "3 Series", "4 Series", "5 Series", "7 Series", "8 Series", "X1", "X2", "X3", "X4", "X5", "X6", "X7", "M2", "M3", "M4", "M5"] },
  { make: "Mercedes-Benz", models: ["A-Class", "B-Class", "C-Class", "E-Class", "S-Class", "CLA", "CLS", "GLA", "GLB", "GLC", "GLE", "GLS", "G-Class", "AMG GT"] },
  { make: "Ford", models: ["Ranger", "Everest", "Fiesta", "Focus", "EcoSport", "Territory", "Mustang", "Edge", "Explorer"] },
  { make: "Hyundai", models: ["i20", "i30", "Tucson", "Creta", "ix35", "Santa Fe", "Elantra", "Venue", "Staria", "H-1", "Grand i10"] },
  { make: "Kia", models: ["Sportage", "Seltos", "Sonet", "Cerato", "Picanto", "Stinger", "Sorento", "Carnival", "Telluride", "EV6"] },
  { make: "Audi", models: ["A1", "A3", "A4", "A5", "A6", "A7", "Q2", "Q3", "Q5", "Q7", "Q8", "TT", "e-tron"] },
  { make: "Nissan", models: ["NP300 Hardbody", "Navara", "X-Trail", "Qashqai", "Micra", "Magnite", "Almera", "Patrol", "Leaf"] },
  { make: "Mazda", models: ["Mazda2", "Mazda3", "CX-3", "CX-30", "CX-5", "CX-60", "BT-50"] },
  { make: "Honda", models: ["Jazz", "Ballade", "Civic", "HR-V", "CR-V", "Accord", "WR-V", "Elevate"] },
  { make: "Suzuki", models: ["Swift", "Jimny", "Vitara", "Grand Vitara", "Baleno", "Ertiga", "S-Presso", "Fronx", "Brezza"] },
  { make: "Renault", models: ["Kwid", "Sandero", "Duster", "Clio", "Captur", "Koleos", "Triber", "Kiger"] },
  { make: "Isuzu", models: ["D-Max", "mu-X", "KB"] },
  { make: "Haval", models: ["H1", "H2", "H6", "Jolion", "F7", "F7x"] },
  { make: "Land Rover", models: ["Range Rover", "Range Rover Sport", "Range Rover Evoque", "Range Rover Velar", "Discovery", "Discovery Sport", "Defender", "Freelander"] },
  { make: "Mitsubishi", models: ["Pajero", "Pajero Sport", "Outlander", "Eclipse Cross", "Triton", "Colt", "ASX"] },
  { make: "Jeep", models: ["Wrangler", "Grand Cherokee", "Cherokee", "Compass", "Renegade", "Gladiator"] },
  { make: "Subaru", models: ["Outback", "Forester", "XV", "Impreza", "Legacy", "WRX", "BRZ"] },
  { make: "Porsche", models: ["911", "Cayenne", "Macan", "Panamera", "Taycan", "718"] },
  { make: "Peugeot", models: ["208", "2008", "308", "3008", "408", "508", "Landtrek"] },
  { make: "Opel", models: ["Corsa", "Astra", "Crossland", "Grandland", "Mokka"] },
  { make: "Mini", models: ["Cooper", "Countryman", "Clubman", "Convertible", "Paceman"] },
  { make: "Chery", models: ["Tiggo 4", "Tiggo 7", "Tiggo 8", "Arrizo 5", "Omoda 5"] },
  { make: "GWM", models: ["P-Series", "Steed 6", "Cannon"] },
  { make: "Chevrolet", models: ["Spark", "Cruze", "Captiva", "Trailblazer", "Utility"] },
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
