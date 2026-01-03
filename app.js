// Farming CBA Decision Tool 2 - Newcastle Business School
// app.js (commercial-grade, Excel-first, control-vs-treatment comparison, export-ready)

(() => {
  "use strict";

  // =========================
  // 0) PRODUCT NAME + VERSION
  // =========================
  const TOOL_NAME = "Farming CBA Decision Tool 2";
  const TOOL_SUBTITLE = "Newcastle Business School";
  const TOOL_VERSION = "2.0.0";

  // =========================
  // 1) DEFAULT DATA (DO NOT OMIT)
  // =========================
  // Stored as raw TSV-like text for reproducibility and Excel sample export.
  // NOTE: Some cells contain '?' exactly as provided.
  const DEFAULT_FABA_BEANS_RAW = String.raw`
Plot	Trt	Rep	Amendment	Practice Change	Plot Length (m)	Plot Width (m)	Plot Area (m^2)	Plants/1m^2	Yield t/ha	Moisture	Protein	Anthesis Biomass t/ha	Harvest Biomass t/ha	Practice Change	Application rate	Treatment Input Cost Only /Ha	Labour per Ha application could be included in next column	Prototype Machinery for Adding amendments	500hp tractor + Speed tiller task	Tractor and 12 m air-seeder wet hire	Sowing Labour included in wet hire	Amberly Faba Bean	Amberly Faba Bean	DAP Fertiliser treated	Inoculant F Pea/Faba	 4Farmers Ammonium Sulphate Herbicide Adjuvant	 4Farmers Ammonium Sulphate Herbicide Adjuvant	Cavalier (Oxyfluofen 240)	Factor	Roundup CT	Roundup Ultra Max	Supercharge Elite Discontinued	Platnium (Clethodim 360)	Mentor	Simazine 900	Veritas Opti	FLUTRIAFOL fungicide	Barrack fungicide discontinued	Barrack fungicide discontinued	Talstar	Talstar	Pre sowing Labour	Amendment Labour	Sowing Labour	Herbicide Labour	Herbicide Labour	Herbicide Labour	Harvesting Labour	Harvesting Labour	Pre sowing amendment 5 tyne ripper	Speed tiller 10 m	Air seeder 12 m	36 m Boomspray	Smaller tractor 150 hp	 Large Tractor 500hp	Header 12 m front	Ute	Truck	Utes $ per kilometer	Trucks $ per kilometer	Tractor	Speed tiller	Air seeder	Boom spray	Header	Truck	|
1	12	1	Deep OM (CP1) + liq. Gypsum (CHT)	Crop 1	20	2.5	50	34	7.03	11.8	23.2	8.40	15.51	Crop 1	15 t/ha ; 0.5 t/ha	$16,850.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00	$210.00	$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$4.24	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$17,945
2	3	1	Deep OM (CP1)	Crop 2	20	2.5	50	27	5.18	10.6	23.6	14.83	16.46	Crop 2	15 t/ha	$16,500.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$17,385
3	11	1	Deep Ripping	Crop 3	20	2.5	50	33	7.26	10.7	23.4	17.89	16.41	Crop 3	n/a	$0.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885
4	1	1	Control	Crop 4	20	2.5	50	29	6.20	10	22.7	12.28	15.19	Crop 4	n/a	$0.00	$0.00	$0.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$0.00	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$695
5	5	1	Deep Carbon-coated mineral (CCM)	Crop 5	20	2.5	50	28	6.13	10.2	22.8	12.69	13.28	Crop 5	5 t/ha	$3,225.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$4,110
6	10	1	Deep OM (CP1) + PAM	Crop 6	20	2.5	50	28	7.27	11.6	23.4	16.13	15.20	Crop 6	15 t/ha ; 5 t/ha	?	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885
7	9	1	Surface Silicon	Crop 7	20	2.5	50	29	6.78	10.5	23.4	12.23	15.29	Crop 7	2 t/ha	?	$35.71	$100.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$835
8	4	1	Deep OM + Gypsum (CP2)	Crop 8	20	2.5	50	31	7.60	10.3	25.2	13.87	14.46	Crop 8	15 t/ha ; 5 t/ha	$24,000.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$24,885
9	6	1	Deep OM (CP1) + Carbon-coated mineral (CCM)	Crop 9	20	2.5	50	31	5.88	10.3	24.4	14.19	17.95	Crop 9	15 t/ha ; 5 t/ha	$21,225.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$22,110
10	7	1	Deep liq. NPKS	Crop 10	20	2.5	50	25	7.23	11.5	23.2	12.12	15.57	Crop 10	750 L/ha	?	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885
11	2	1	Deep Gypsum	Crop 11	20	2.5	50	22	6.29	9.9	22.8	9.85	15.45	Crop 11	5 t/ha	$500.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$1,385
12	8	1	Deep liq. Gypsum (CHT)	Crop 12	20	2.5	50	26	5.88	9.9	23.5	10.48	11.69	Crop 12	0.5 t/ha	$350.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$1,235
13	6	2	Deep OM (CP1) + Carbon-coated mineral (CCM)	Crop 13	20	2.5	50	33	4.79	9.8	24.4	14.49	13.62	Crop 13	15 t/ha ; 5 t/ha	$21,225.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$22,110
14	7	2	Deep liq. NPKS	Crop 14	20	2.5	50	29	4.88	10.4	23.7	12.81	13.49	Crop 14	750 L/ha	?	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885
15	5	2	Deep Carbon-coated mineral (CCM)	Crop 15	20	2.5	50	26	5.39	10.5	23.7	11.97	12.77	Crop 15	5 t/ha	$3,225.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$4,110
16	3	2	Deep OM (CP1)	Crop 16	20	2.5	50	24	4.96	10.2	23.2	13.85	14.44	Crop 16	15 t/ha	$16,500.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$17,385
17	1	2	Control	Crop 17	20	2.5	50	24	4.99	10.3	23.3	15.61	10.63	Crop 17	n/a	$0.00	$0.00	$0.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$0.00	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$695
18	9	2	Surface Silicon	Crop 18	20	2.5	50	27	5.79	10.6	21.1	8.59	10.63	Crop 18	2 t/ha	?	$35.71	$100.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$835
19	4	2	Deep OM + Gypsum (CP2)	Crop 19	20	2.5	50	22	5.45	11.2	23	12.34	15.59	Crop 19	15 t/ha ; 5 t/ha	$24,000.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$24,885
20	11	2	Deep Ripping	Crop 20	20	2.5	50	27	6.30	10.4	22.9	12.34	15.28	Crop 20	n/a	$0.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885
21	8	2	Deep liq. Gypsum (CHT)	Crop 21	20	2.5	50	24	6.57	9.8	23.2	16.16	11.35	Crop 21	0.5 t/ha	$350.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$1,235
22	12	2	Deep OM (CP1) + liq. Gypsum (CHT)	Crop 22	20	2.5	50	26	6.10	10.3	23.6	14.16	12.21	Crop 22	15 t/ha ; 0.5 t/ha	$16,850.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$17,735
23	10	2	Deep OM (CP1) + PAM	Crop 23	20	2.5	50	24	6.34	10	22.8	15.68	12.70	Crop 23	15 t/ha ; 5 t/ha	?	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885
24	2	2	Deep Gypsum	Crop 24	20	2.5	50	25	5.44	9.8	23.1	12.70	13.24	Crop 24	5 t/ha	$500.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$1,385
25	6	3	Deep OM (CP1) + Carbon-coated mineral (CCM)	Crop 25	20	2.5	50	19	5.04	11.2	24.5	15.45	10.97	Crop 25	15 t/ha ; 5 t/ha	$21,225.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$22,110
26	11	3	Deep Ripping	Crop 26	20	2.5	50	21	6.35	11.2	27.3	19.73	20.65	Crop 26	n/a	$0.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885
27	2	3	Deep Gypsum	Crop 27	20	2.5	50	21	6.94	10.2	24.6	16.39	14.52	Crop 27	5 t/ha	$500.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$1,385
28	5	3	Deep Carbon-coated mineral (CCM)	Crop 28	20	2.5	50	19	6.31	10.2	23	11.23	15.58	Crop 28	5 t/ha	$3,225.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$4,110
29	8	3	Deep liq. Gypsum (CHT)	Crop 29	20	2.5	50	26	6.64	11.2	23.5	13.36	14.23	Crop 29	0.5 t/ha	$350.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$1,235
30	12	3	Deep OM (CP1) + liq. Gypsum (CHT)	Crop 30	20	2.5	50	22	5.96	10.4	23.8	12.01	13.71	Crop 30	15 t/ha ; 0.5 t/ha	$16,850.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$17,735
31	10	3	Deep OM (CP1) + PAM	Crop 31	20	2.5	50	22	7.58	10.2	24.2	12.73	11.98	Crop 31	15 t/ha ; 5 t/ha	?	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885
32	4	3	Deep OM + Gypsum (CP2)	Crop 32	20	2.5	50	25	6.68	10.3	24.6	13.34	13.12	Crop 32	15 t/ha ; 5 t/ha	$24,000.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$24,885
33	7	3	Deep liq. NPKS	Crop 33	20	2.5	50	23	7.33	10.1	23.3	13.06	12.18	Crop 33	750 L/ha	?	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885
34	1	3	Control	Crop 34	20	2.5	50	25	7.37	10.3	23.3	15.30	9.52	Crop 34	n/a	$0.00	$0.00	$0.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$0.00	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$695
35	3	3	Deep OM (CP1)	Crop 35	20	2.5	50	23	5.29	10.5	23.7	12.61	11.73	Crop 35	15 t/ha	$16,500.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$17,385
36	9	3	Surface Silicon	Crop 36	20	2.5	50	18	6.81	10	23.8	14.04	17.68	Crop 36	2 t/ha	?	$35.71	$100.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$835
37	5	4	Deep Carbon-coated mineral (CCM)	Crop 37	20	2.5	50	20	6.42	11.1	23.4	13.51	13.34	Crop 37	5 t/ha	$3,225.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$4,110
38	6	4	Deep OM (CP1) + Carbon-coated mineral (CCM)	Crop 38	20	2.5	50	20	6.18	10.6	24.9	14.50	13.16	Crop 38	15 t/ha ; 5 t/ha	$21,225.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$22,110
39	9	4	Surface Silicon	Crop 39	20	2.5	50	21	6.69	10.8	24.6	13.72	15.00	Crop 39	2 t/ha	?	$35.71	$100.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$835
40	10	4	Deep OM (CP1) + PAM	Crop 40	20	2.5	50	21	7.72	10.2	23.3	16.55	18.02	Crop 40	15 t/ha ; 5 t/ha	?	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885
41	11	4	Deep Ripping	Crop 41	20	2.5	50	23	6.28	10.6	23.4	10.25	14.71	Crop 41	n/a	$0.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885
42	2	4	Deep Gypsum	Crop 42	20	2.5	50	19	5.85	9.8	23.1	10.66	11.19	Crop 42	5 t/ha	$500.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$1,385
43	7	4	Deep liq. NPKS	Crop 43	20	2.5	50	23	6.40	10.1	23.6	13.28	10.18	Crop 43	750 L/ha	?	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885
44	4	4	Deep OM + Gypsum (CP2)	Crop 44	20	2.5	50	33	5.30	9.7	25.5	16.80	13.87	Crop 44	15 t/ha ; 5 t/ha	$24,000.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$24,885
45	1	4	Control	Crop 45	20	2.5	50	24	6.21	9.9	22.1	10.02	14.31	Crop 45	n/a	$0.00	$0.00	$0.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$0.00	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$695
46	3	4	Deep OM (CP1)	Crop 46	20	2.5	50	28	5.85	10.9	23.9	13.05	13.28	Crop 46	15 t/ha	$16,500.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$17,385
47	8	4	Deep liq. Gypsum (CHT)	Crop 47	20	2.5	50	27	5.85	9.6	24.2	20.66	12.83	Crop 47	0.5 t/ha	$350.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$1,235
48	12	4	Deep OM (CP1) + liq. Gypsum (CHT)	Crop 48	20	2.5	50	23	6.06	10	25.1	15.65	11.32	Crop 48	15 t/ha ; 0.5 t/ha	$16,850.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$17,735
`.trim();

  // =========================
  // 2) DEFAULT SETTINGS
  // =========================
  const DEFAULT_DISCOUNT_SCHEDULE = [
    { label: "2025-2034", from: 2025, to: 2034, low: 2, base: 4, high: 6 },
    { label: "2035-2044", from: 2035, to: 2044, low: 4, base: 7, high: 10 },
    { label: "2045-2054", from: 2045, to: 2054, low: 4, base: 7, high: 10 },
    { label: "2055-2064", from: 2055, to: 2064, low: 3, base: 6, high: 9 },
    { label: "2065-2074", from: 2065, to: 2074, low: 2, base: 5, high: 8 }
  ];
  const horizons = [5, 10, 15, 20, 25];

  // =========================
  // 3) SMALL HELPERS
  // =========================
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const esc = s =>
    (s ?? "")
      .toString()
      .replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  function uid() {
    return Math.random().toString(36).slice(2, 10);
  }

  function toNum(x) {
    if (x === null || x === undefined) return NaN;
    if (typeof x === "number") return x;
    const s = String(x).trim();
    if (!s) return NaN;
    // Keep '?' as NaN by design (explicit missing / to be filled)
    if (s === "?") return NaN;
    const cleaned = s
      .replace(/\$/g, "")
      .replace(/,/g, "")
      .replace(/\s+/g, "")
      .replace(/%/g, "");
    const v = Number(cleaned);
    return isFinite(v) ? v : NaN;
  }

  function fmt(n, maxFrac = 2) {
    if (!isFinite(n)) return "n/a";
    const abs = Math.abs(n);
    if (abs >= 1000) return n.toLocaleString(undefined, { maximumFractionDigits: 0 });
    return n.toLocaleString(undefined, { maximumFractionDigits: maxFrac });
  }

  function money(n) {
    return isFinite(n) ? "$" + fmt(n, 2) : "n/a";
  }

  function percent(n) {
    return isFinite(n) ? fmt(n, 2) + "%" : "n/a";
  }

  function slug(s) {
    return (s || "project")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_|_$/g, "");
  }

  function showToast(message) {
    const root = document.getElementById("toast-root") || document.body;
    const toast = document.createElement("div");
    toast.className = "toast";
    toast.textContent = message;
    root.appendChild(toast);
    void toast.offsetWidth;
    toast.classList.add("show");
    setTimeout(() => {
      toast.classList.remove("show");
      setTimeout(() => toast.remove(), 200);
    }, 3500);
  }

  function ensureXLSX() {
    if (!window.XLSX) {
      showToast("Excel features require the XLSX library. Please check your internet connection.");
      return false;
    }
    return true;
  }

  function rng(seed) {
    let t = (seed || Math.floor(Math.random() * 2 ** 31)) >>> 0;
    return () => {
      t += 0x6d2b79f5;
      let x = t;
      x = Math.imul(x ^ (x >>> 15), 1 | x);
      x ^= x + Math.imul(x ^ (x >>> 7), 61 | x);
      return ((x ^ (x >>> 14)) >>> 0) / 4294967296;
    };
  }

  function triangular(r, a, c, b) {
    const F = (c - a) / (b - a);
    if (r < F) return a + Math.sqrt(r * (b - a) * (c - a));
    return b - Math.sqrt((1 - r) * (b - a) * (b - c));
  }

  // =========================
  // 4) MODEL
  // =========================
  const model = {
    meta: {
      toolName: TOOL_NAME,
      version: TOOL_VERSION
    },
    project: {
      name: "Faba bean soil amendment trial",
      lead: "Project lead",
      analysts: "Farm economics team",
      team: "Trial team",
      organisation: "Newcastle Business School, The University of Newcastle",
      contactEmail: "",
      contactPhone: "",
      summary:
        "Applied faba bean trial comparing deep ripping, organic matter, gypsum and fertiliser treatments against a control.",
      objectives: "Quantify yield and gross margin impacts of alternative soil amendment strategies.",
      activities:
        "Establish replicated field plots, collect plot-level yield and cost data, and summarise trial-wide economics.",
      stakeholders: "Producers, agronomists, government agencies, research partners.",
      lastUpdated: new Date().toISOString().slice(0, 10),
      goal:
        "Identify soil amendment packages that deliver higher faba bean yields and acceptable returns after accounting for additional costs.",
      withProject:
        "Growers adopt high-performing amendment packages on trial farms and similar soils in the region.",
      withoutProject:
        "Growers continue with baseline practice and do not access detailed economic evidence on soil amendments."
    },
    time: {
      startYear: new Date().getFullYear(),
      projectStartYear: new Date().getFullYear(),
      years: 10,
      discBase: 7,
      discLow: 4,
      discHigh: 10,
      mirrFinance: 6,
      mirrReinvest: 4,
      discountSchedule: JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE))
    },
    outputsMeta: {
      systemType: "single",
      assumptions: ""
    },
    outputs: [
      // Core output: yield
      { id: uid(), name: "Grain yield", unit: "t/ha", value: 450, source: "Default (example)" }
    ],
    treatments: [],
    benefits: [],
    otherCosts: [],
    adoption: { base: 0.9, low: 0.6, high: 1.0 },
    risk: {
      base: 0.15,
      low: 0.05,
      high: 0.3,
      tech: 0.05,
      nonCoop: 0.04,
      socio: 0.02,
      fin: 0.03,
      man: 0.02
    },
    sim: {
      n: 1000,
      targetBCR: 2,
      bcrMode: "all",
      seed: null,
      results: { npv: [], bcr: [] },
      details: [],
      variationPct: 20,
      varyOutputs: true,
      varyTreatCosts: true,
      varyInputCosts: false
    },
    raw: {
      fabaBeans2022: DEFAULT_FABA_BEANS_RAW
    },
    computed: {
      lastRun: null,
      perTreatment: [],
      totals: null,
      compareTable: null,
      timeProjection: []
    }
  };

  let parsedExcel = null;

  // =========================
  // 5) DEFAULT CALIBRATION FROM RAW DATA
  // =========================
  function parseRawTSV(tsv) {
    const lines = String(tsv || "")
      .split(/\r?\n/)
      .filter(l => l.trim().length > 0);

    if (lines.length < 2) return { headers: [], rows: [] };

    const headers = lines[0].split("\t").map(h => h.trim());
    const rows = lines.slice(1).map(line => {
      const cols = line.split("\t");
      const row = {};
      headers.forEach((h, i) => (row[h] = cols[i] ?? ""));
      return row;
    });

    return { headers, rows };
  }

  function summariseTreatmentsFromFabaRaw(tsv) {
    const { rows } = parseRawTSV(tsv);

    // We build treatment groups by "Amendment" (text label), and keep the "Control" as explicit control.
    // Yields are averaged within each amendment group.
    const groups = new Map();

    for (const r of rows) {
      const amend = (r["Amendment"] || "").trim() || "Unknown treatment";
      const yieldTHa = toNum(r["Yield t/ha"]);
      const inputCostOnly = toNum(r["Treatment Input Cost Only /Ha"]);

      if (!groups.has(amend)) {
        groups.set(amend, {
          amendment: amend,
          n: 0,
          yieldSum: 0,
          yieldN: 0,
          inputCostOnlySum: 0,
          inputCostOnlyN: 0,
          // keep any representative fields
          trtCodes: new Set(),
          practiceChanges: new Set()
        });
      }
      const g = groups.get(amend);
      g.n += 1;
      const trt = (r["Trt"] || "").toString().trim();
      if (trt) g.trtCodes.add(trt);
      const pc = (r["Practice Change"] || "").toString().trim();
      if (pc) g.practiceChanges.add(pc);

      if (isFinite(yieldTHa)) {
        g.yieldSum += yieldTHa;
        g.yieldN += 1;
      }
      if (isFinite(inputCostOnly)) {
        g.inputCostOnlySum += inputCostOnly;
        g.inputCostOnlyN += 1;
      }
    }

    // Identify control group(s)
    // Provided data uses Amendment = "Control"
    let controlYield = NaN;
    let controlCostOnly = NaN;
    if (groups.has("Control")) {
      const c = groups.get("Control");
      controlYield = c.yieldN ? c.yieldSum / c.yieldN : NaN;
      controlCostOnly = c.inputCostOnlyN ? c.inputCostOnlySum / c.inputCostOnlyN : NaN;
    }

    // Build treatment objects
    const outputId = model.outputs[0]?.id;

    const treatments = [];
    const amendmentsSorted = Array.from(groups.values()).sort((a, b) => a.amendment.localeCompare(b.amendment));

    for (const g of amendmentsSorted) {
      const name = g.amendment;
      const avgYield = g.yieldN ? g.yieldSum / g.yieldN : NaN;
      const avgInputCost = g.inputCostOnlyN ? g.inputCostOnlySum / g.inputCostOnlyN : NaN;

      const isControl = name.toLowerCase() === "control";

      // delta yield is relative to control yield (if control exists)
      const deltaYield = isControl ? 0 : isFinite(avgYield) && isFinite(controlYield) ? avgYield - controlYield : 0;

      treatments.push({
        id: uid(),
        name,
        area: 100, // default example: 100 ha paddock (as per text)
        adoption: 1,
        isControl,
        constrained: true,
        notes: [
          g.trtCodes.size ? `Trt code(s): ${Array.from(g.trtCodes).join(", ")}` : "",
          g.practiceChanges.size ? `Practice change(s): ${Array.from(g.practiceChanges).join(", ")}` : ""
        ]
          .filter(Boolean)
          .join(" | "),
        source: "Default dataset (Faba beans 2022)",
        // outputs deltas
        deltas: outputId ? { [outputId]: isFinite(deltaYield) ? deltaYield : 0 } : {},
        // Cost structure (per ha unless stated)
        // Capital cost is year 0 (explicit field)
        capitalCostYear0: 0,
        // cost components per ha (applied in year 0 by default)
        costComponents: {
          inputCostOnlyPerHa: isFinite(avgInputCost) ? avgInputCost : NaN,
          labourPerHa: NaN,
          servicesPerHa: NaN,
          otherPerHa: NaN
        },
        costYears: 1, // default: apply treatment costs in year 0 only
        benefitYears: model.time.years // default: yield benefit across full horizon (user can adjust)
      });
    }

    // Ensure there is exactly one control flagged: if "Control" exists, use it; otherwise create one.
    const hasControl = treatments.some(t => t.isControl);
    if (!hasControl) {
      treatments.unshift({
        id: uid(),
        name: "Control",
        area: 100,
        adoption: 1,
        isControl: true,
        constrained: true,
        notes: "Baseline practice (control).",
        source: "Default dataset (constructed)",
        deltas: outputId ? { [outputId]: 0 } : {},
        capitalCostYear0: 0,
        costComponents: {
          inputCostOnlyPerHa: 0,
          labourPerHa: 0,
          servicesPerHa: 0,
          otherPerHa: 0
        },
        costYears: 1,
        benefitYears: model.time.years
      });
    }

    // If control has missing cost-only, set to 0 (conservative baseline)
    const control = treatments.find(t => t.isControl);
    if (control && !isFinite(control.costComponents.inputCostOnlyPerHa)) {
      control.costComponents.inputCostOnlyPerHa = 0;
    }

    return treatments;
  }

  function loadDefaultScenario() {
    model.treatments = summariseTreatmentsFromFabaRaw(model.raw.fabaBeans2022);

    // Minimal example: no additional benefits/costs by default; keep tool clean.
    model.benefits = [];
    model.otherCosts = [];

    // Initialise deltas for any outputs
    initTreatmentDeltas();
  }

  function initTreatmentDeltas() {
    const outputs = model.outputs || [];
    for (const t of model.treatments) {
      if (!t.deltas) t.deltas = {};
      for (const o of outputs) {
        if (!(o.id in t.deltas)) t.deltas[o.id] = 0;
      }
      // Ensure cost components exist
      if (!t.costComponents) {
        t.costComponents = { inputCostOnlyPerHa: 0, labourPerHa: 0, servicesPerHa: 0, otherPerHa: 0 };
      } else {
        t.costComponents.inputCostOnlyPerHa = t.costComponents.inputCostOnlyPerHa ?? 0;
        t.costComponents.labourPerHa = t.costComponents.labourPerHa ?? 0;
        t.costComponents.servicesPerHa = t.costComponents.servicesPerHa ?? 0;
        t.costComponents.otherPerHa = t.costComponents.otherPerHa ?? 0;
      }
      t.capitalCostYear0 = t.capitalCostYear0 ?? 0;
      t.costYears = t.costYears ?? 1;
      t.benefitYears = t.benefitYears ?? model.time.years;
      t.area = isFinite(toNum(t.area)) ? Number(t.area) : 0;
      t.adoption = isFinite(toNum(t.adoption)) ? Number(t.adoption) : 1;
      t.constrained = !!t.constrained;
      t.isControl = !!t.isControl;
    }
  }

  // =========================
  // 6) CBA CORE CALCULATIONS
  // =========================
  function discountFactor(yearIndex, discPct) {
    const r = discPct / 100;
    return 1 / Math.pow(1 + r, yearIndex);
  }

  function calcPerTreatmentCashflows(t, opts) {
    const years = opts.years;
    const disc = opts.discBase;
    const risk = clamp(opts.riskBase ?? 0, 0, 1);
    const adoptionMult = clamp(opts.adoptBase ?? 1, 0, 1);

    const output = model.outputs[0];
    const outputValue = isFinite(output?.value) ? output.value : 0;

    const deltaYield = isFinite(toNum(t.deltas?.[output?.id])) ? Number(t.deltas[output.id]) : 0;
    const area = isFinite(toNum(t.area)) ? Number(t.area) : 0;
    const treatAdopt = clamp(isFinite(toNum(t.adoption)) ? Number(t.adoption) : 1, 0, 1);

    // Effective adoption scaling: scenario adoption * treatment adoption * area
    const scale = adoptionMult * treatAdopt * area;

    // Benefits: annual yield delta * price (value) * area * adoption; reduced by risk
    // Benefits apply from year 1 to benefitYears (default full horizon). Year index is 0-based.
    const benefitYears = clamp(Math.round(toNum(t.benefitYears) || years), 0, years);
    const annualBenefit = deltaYield * outputValue * scale * (1 - risk);

    // Costs: per ha cost components apply for costYears starting year 0
    const costYears = clamp(Math.round(toNum(t.costYears) || 1), 0, years);

    const cc = t.costComponents || {};
    const inputCostOnlyPerHa = isFinite(toNum(cc.inputCostOnlyPerHa)) ? Number(cc.inputCostOnlyPerHa) : 0;
    const labourPerHa = isFinite(toNum(cc.labourPerHa)) ? Number(cc.labourPerHa) : 0;
    const servicesPerHa = isFinite(toNum(cc.servicesPerHa)) ? Number(cc.servicesPerHa) : 0;
    const otherPerHa = isFinite(toNum(cc.otherPerHa)) ? Number(cc.otherPerHa) : 0;

    const totalVarCostPerHa = inputCostOnlyPerHa + labourPerHa + servicesPerHa + otherPerHa;
    const annualVarCost = totalVarCostPerHa * scale;

    const capitalYear0 = isFinite(toNum(t.capitalCostYear0)) ? Number(t.capitalCostYear0) : 0;

    const benefits = new Array(years + 1).fill(0);
    const costs = new Array(years + 1).fill(0);

    // Year 0 costs include capital + first-year variable cost (if costYears >= 1)
    costs[0] += capitalYear0;
    if (costYears >= 1) costs[0] += annualVarCost;

    // Remaining cost years: years 1..(costYears-1)
    for (let y = 1; y < costYears; y++) costs[y] += annualVarCost;

    // Benefits: years 1..benefitYears (if benefitYears is count)
    for (let y = 1; y <= benefitYears; y++) benefits[y] += annualBenefit;

    // Discounted PV
    let pvB = 0;
    let pvC = 0;
    const discBenefits = [];
    const discCosts = [];
    const net = [];
    const discNet = [];

    for (let y = 0; y <= years; y++) {
      const df = discountFactor(y, disc);
      const db = benefits[y] * df;
      const dc = costs[y] * df;
      pvB += db;
      pvC += dc;
      discBenefits.push(db);
      discCosts.push(dc);
      net.push(benefits[y] - costs[y]);
      discNet.push((benefits[y] - costs[y]) * df);
    }

    const npv = pvB - pvC;
    const bcr = pvC !== 0 ? pvB / pvC : NaN;
    const roi = pvC !== 0 ? (npv / pvC) * 100 : NaN;

    return {
      benefits,
      costs,
      net,
      pvBenefits: pvB,
      pvCosts: pvC,
      npv,
      bcr,
      roi,
      totalVarCostPerHa,
      annualBenefit,
      annualVarCost,
      scale
    };
  }

  function calcProjectTotals(perTreatment) {
    // Whole-project totals = sum across treatments
    const pvB = perTreatment.reduce((s, x) => s + (isFinite(x.pvBenefits) ? x.pvBenefits : 0), 0);
    const pvC = perTreatment.reduce((s, x) => s + (isFinite(x.pvCosts) ? x.pvCosts : 0), 0);
    const npv = pvB - pvC;
    const bcr = pvC !== 0 ? pvB / pvC : NaN;
    const roi = pvC !== 0 ? (npv / pvC) * 100 : NaN;
    return { pvBenefits: pvB, pvCosts: pvC, npv, bcr, roi };
  }

  function findControlTreatment() {
    const controls = model.treatments.filter(t => t.isControl);
    if (controls.length === 1) return controls[0];
    if (controls.length > 1) return controls[0]; // deterministic: first
    // fallback: name contains "control"
    const byName = model.treatments.find(t => (t.name || "").toLowerCase().includes("control"));
    return byName || null;
  }

  function computeAll() {
    const years = clamp(Math.round(toNum(model.time.years) || 10), 1, 100);
    const discBase = toNum(model.time.discBase) || 7;
    const riskBase = clamp(toNum(model.risk.base) || 0, 0, 1);
    const adoptBase = clamp(toNum(model.adoption.base) || 1, 0, 1);

    const opts = { years, discBase, riskBase, adoptBase };

    const perTreatment = model.treatments.map(t => ({
      id: t.id,
      name: t.name,
      isControl: t.isControl,
      constrained: t.constrained,
      t,
      ...calcPerTreatmentCashflows(t, opts)
    }));

    // Rankings (snapshot): rank by BCR (descending), tie-break by NPV
    const ranked = perTreatment
      .slice()
      .sort((a, b) => {
        const ab = isFinite(a.bcr) ? a.bcr : -Infinity;
        const bb = isFinite(b.bcr) ? b.bcr : -Infinity;
        if (bb !== ab) return bb - ab;
        const an = isFinite(a.npv) ? a.npv : -Infinity;
        const bn = isFinite(b.npv) ? b.npv : -Infinity;
        return bn - an;
      })
      .map((x, i) => ({ ...x, rank: i + 1 }));

    // Control reference row (for narrative and optional incremental rows)
    const control = ranked.find(x => x.isControl) || ranked.find(x => x.t && x.t.isControl) || null;

    // Build comparison table (rows = indicators; cols = treatments incl control)
    const cols = ranked.map(x => ({
      id: x.id,
      name: x.name,
      isControl: x.isControl,
      rank: x.rank
    }));

    const rows = [
      { key: "pvBenefits", label: "Present value of benefits", fmt: money },
      { key: "pvCosts", label: "Present value of costs", fmt: money },
      { key: "npv", label: "Net present value", fmt: money },
      { key: "bcr", label: "Benefit–cost ratio", fmt: v => (isFinite(v) ? fmt(v, 3) : "n/a") },
      { key: "roi", label: "Return on investment", fmt: v => (isFinite(v) ? percent(v) : "n/a") },
      { key: "rank", label: "Ranking (BCR-based)", fmt: v => (isFinite(v) ? String(v) : "n/a") }
    ];

    // Optional incremental rows vs control (helpful snapshot; does not replace required rows)
    if (control) {
      rows.push(
        { key: "inc_npv", label: "Incremental NPV vs control", fmt: money },
        { key: "inc_pvBenefits", label: "Incremental PV benefits vs control", fmt: money },
        { key: "inc_pvCosts", label: "Incremental PV costs vs control", fmt: money }
      );
    }

    const cell = (tRow, key) => {
      if (key === "rank") return tRow.rank;
      if (key.startsWith("inc_") && control) {
        const baseKey = key.replace(/^inc_/, "");
        const v = tRow[baseKey];
        const c = control[baseKey];
        return isFinite(v) && isFinite(c) ? v - c : NaN;
      }
      return tRow[key];
    };

    const table = {
      rows,
      cols,
      values: rows.map(r =>
        cols.map(c => {
          const tr = ranked.find(x => x.id === c.id);
          return tr ? cell(tr, r.key) : NaN;
        })
      )
    };

    // Whole project totals
    const totals = calcProjectTotals(perTreatment);

    // Time projection (based on whole project totals, by truncating horizon)
    // Uses current per-treatment cashflows to compute partial PVs.
    const timeProjection = horizons
      .filter(h => h <= years)
      .map(h => {
        let pvB = 0;
        let pvC = 0;
        for (const x of perTreatment) {
          for (let y = 0; y <= h; y++) {
            const df = discountFactor(y, discBase);
            pvB += (x.benefits[y] || 0) * df;
            pvC += (x.costs[y] || 0) * df;
          }
        }
        const npv = pvB - pvC;
        const bcr = pvC !== 0 ? pvB / pvC : NaN;
        return { years: h, pvBenefits: pvB, pvCosts: pvC, npv, bcr };
      });

    model.computed.lastRun = new Date().toISOString();
    model.computed.perTreatment = ranked;
    model.computed.totals = totals;
    model.computed.compareTable = table;
    model.computed.timeProjection = timeProjection;

    return model.computed;
  }

  // =========================
  // 7) UI: TABS + BASIC BINDINGS
  // =========================
  function setToolBranding() {
    document.title = TOOL_NAME;
    const brandTitle = $(".brand-title");
    if (brandTitle) brandTitle.textContent = TOOL_NAME;
    const footerText = $(".app-footer .footer-left .small");
    if (footerText) footerText.textContent = `${TOOL_NAME}, ${TOOL_SUBTITLE}`;
    const introH1 = $("#tab-intro h1");
    if (introH1) introH1.textContent = TOOL_NAME;
    const headerTitle = $("head > title");
    if (headerTitle) headerTitle.textContent = TOOL_NAME;
  }

  function activateTab(tabName) {
    const buttons = $$(".tab-link");
    const panels = $$(".tab-panel");

    buttons.forEach(b => {
      const isActive = b.dataset.tab === tabName;
      b.classList.toggle("active", isActive);
      b.setAttribute("aria-selected", isActive ? "true" : "false");
    });

    panels.forEach(p => {
      const isActive = p.dataset.tabPanel === tabName;
      p.classList.toggle("active", isActive);
      p.classList.toggle("show", isActive);
      p.setAttribute("aria-hidden", isActive ? "false" : "true");
    });
  }

  function initTabs() {
    $$(".tab-link").forEach(btn => {
      btn.addEventListener("click", () => activateTab(btn.dataset.tab));
    });

    $$("[data-tab-jump]").forEach(btn => {
      btn.addEventListener("click", () => activateTab(btn.dataset.tabJump));
    });
  }

  // =========================
  // 8) UI: PROJECT + SETTINGS BIND
  // =========================
  function bindProjectFields() {
    const map = [
      ["projectName", "name"],
      ["projectLead", "lead"],
      ["analystNames", "analysts"],
      ["projectTeam", "team"],
      ["organisation", "organisation"],
      ["lastUpdated", "lastUpdated"],
      ["contactEmail", "contactEmail"],
      ["contactPhone", "contactPhone"],
      ["projectSummary", "summary"],
      ["projectGoal", "goal"],
      ["withProject", "withProject"],
      ["withoutProject", "withoutProject"],
      ["projectObjectives", "objectives"],
      ["projectActivities", "activities"],
      ["stakeholderGroups", "stakeholders"]
    ];

    map.forEach(([id, key]) => {
      const el = $("#" + id);
      if (!el) return;
      el.value = model.project[key] ?? "";
      el.addEventListener("input", () => {
        model.project[key] = el.value;
        refreshCopilotPreview();
      });
    });
  }

  function bindSettingsFields() {
    const ids = [
      ["startYear", "startYear"],
      ["projectStartYear", "projectStartYear"],
      ["years", "years"],
      ["discBase", "discBase"],
      ["discLow", "discLow"],
      ["discHigh", "discHigh"],
      ["mirrFinance", "mirrFinance"],
      ["mirrReinvest", "mirrReinvest"]
    ];

    ids.forEach(([id, key]) => {
      const el = $("#" + id);
      if (!el) return;
      el.value = model.time[key] ?? "";
      el.addEventListener("input", () => {
        model.time[key] = toNum(el.value);
        refreshAll();
      });
    });

    const systemType = $("#systemType");
    if (systemType) {
      systemType.value = model.outputsMeta.systemType || "single";
      systemType.addEventListener("change", () => {
        model.outputsMeta.systemType = systemType.value;
        refreshCopilotPreview();
      });
    }

    const assumptions = $("#outputAssumptions");
    if (assumptions) {
      assumptions.value = model.outputsMeta.assumptions || "";
      assumptions.addEventListener("input", () => {
        model.outputsMeta.assumptions = assumptions.value;
        refreshCopilotPreview();
      });
    }

    // Adoption
    const adoptLow = $("#adoptLow");
    const adoptBase = $("#adoptBase");
    const adoptHigh = $("#adoptHigh");
    if (adoptLow) {
      adoptLow.value = model.adoption.low;
      adoptLow.addEventListener("input", () => {
        model.adoption.low = clamp(toNum(adoptLow.value), 0, 1);
        refreshAll();
      });
    }
    if (adoptBase) {
      adoptBase.value = model.adoption.base;
      adoptBase.addEventListener("input", () => {
        model.adoption.base = clamp(toNum(adoptBase.value), 0, 1);
        refreshAll();
      });
    }
    if (adoptHigh) {
      adoptHigh.value = model.adoption.high;
      adoptHigh.addEventListener("input", () => {
        model.adoption.high = clamp(toNum(adoptHigh.value), 0, 1);
        refreshAll();
      });
    }

    // Risk
    const riskLow = $("#riskLow");
    const riskBase = $("#riskBase");
    const riskHigh = $("#riskHigh");
    const rTech = $("#rTech");
    const rNonCoop = $("#rNonCoop");
    const rSocio = $("#rSocio");
    const rFin = $("#rFin");
    const rMan = $("#rMan");

    const bindRisk = (el, key) => {
      if (!el) return;
      el.value = model.risk[key] ?? 0;
      el.addEventListener("input", () => {
        model.risk[key] = clamp(toNum(el.value), 0, 1);
        refreshAll();
      });
    };

    bindRisk(riskLow, "low");
    bindRisk(riskBase, "base");
    bindRisk(riskHigh, "high");
    bindRisk(rTech, "tech");
    bindRisk(rNonCoop, "nonCoop");
    bindRisk(rSocio, "socio");
    bindRisk(rFin, "fin");
    bindRisk(rMan, "man");

    const calcBtn = $("#calcCombinedRisk");
    const combinedOut = $("#combinedRiskOut .value");
    if (calcBtn) {
      calcBtn.addEventListener("click", () => {
        // combined risk = 1 - Π(1 - r_i)
        const rs = ["tech", "nonCoop", "socio", "fin", "man"].map(k => clamp(toNum(model.risk[k]) || 0, 0, 1));
        const combined = 1 - rs.reduce((p, r) => p * (1 - r), 1);
        model.risk.base = clamp(combined, 0, 1);
        if (riskBase) riskBase.value = model.risk.base;
        if (combinedOut) combinedOut.textContent = fmt(model.risk.base, 3);
        showToast("Combined risk updated into base risk.");
        refreshAll();
      });
    }

    // Discount schedule inputs
    $$("[data-disc-period]").forEach(input => {
      const periodIndex = Number(input.dataset.discPeriod);
      const scenario = input.dataset.scenario;
      const row = model.time.discountSchedule[periodIndex];
      if (!row) return;
      input.value = row[scenario] ?? "";
      input.addEventListener("input", () => {
        row[scenario] = toNum(input.value);
        refreshAll();
      });
    });
  }

  // =========================
  // 9) UI: OUTPUTS
  // =========================
  function renderOutputs() {
    const list = $("#outputsList");
    if (!list) return;
    list.innerHTML = "";

    model.outputs.forEach(o => {
      const card = document.createElement("div");
      card.className = "item";
      card.innerHTML = `
        <div class="item-head">
          <div class="item-title">${esc(o.name)}</div>
          <div class="item-actions">
            <button class="btn small ghost" data-act="del">Remove</button>
          </div>
        </div>
        <div class="row-4">
          <div class="field">
            <label data-tooltip="Output label used in benefits calculations.">Output name</label>
            <input type="text" value="${esc(o.name)}" data-k="name"/>
          </div>
          <div class="field">
            <label data-tooltip="Unit used for the output (for example t/ha).">Unit</label>
            <input type="text" value="${esc(o.unit)}" data-k="unit"/>
          </div>
          <div class="field">
            <label data-tooltip="Monetary value per unit used to convert output change into benefits.">Value per unit (AUD)</label>
            <input type="number" step="0.01" value="${isFinite(toNum(o.value)) ? Number(o.value) : ""}" data-k="value"/>
          </div>
          <div class="field">
            <label data-tooltip="Optional source or note for this output value.">Source</label>
            <input type="text" value="${esc(o.source || "")}" data-k="source"/>
          </div>
        </div>
      `;

      const del = $('[data-act="del"]', card);
      del.addEventListener("click", () => {
        model.outputs = model.outputs.filter(x => x.id !== o.id);
        initTreatmentDeltas();
        renderOutputs();
        renderTreatments();
        refreshAll();
      });

      $$("input[data-k]", card).forEach(inp => {
        inp.addEventListener("input", () => {
          const k = inp.dataset.k;
          if (k === "value") o[k] = toNum(inp.value);
          else o[k] = inp.value;
          initTreatmentDeltas();
          refreshAll();
        });
      });

      list.appendChild(card);
    });
  }

  function bindOutputsButtons() {
    const addBtn = $("#addOutput");
    if (!addBtn) return;
    addBtn.addEventListener("click", () => {
      model.outputs.push({ id: uid(), name: "New output", unit: "unit", value: 0, source: "" });
      initTreatmentDeltas();
      renderOutputs();
      renderTreatments();
      refreshAll();
    });
  }

  // =========================
  // 10) UI: TREATMENTS (capital cost before total cost)
  // =========================
  function computeTreatmentTotalCostPerHa(t) {
    const cc = t.costComponents || {};
    const a = isFinite(toNum(cc.inputCostOnlyPerHa)) ? Number(cc.inputCostOnlyPerHa) : 0;
    const b = isFinite(toNum(cc.labourPerHa)) ? Number(cc.labourPerHa) : 0;
    const c = isFinite(toNum(cc.servicesPerHa)) ? Number(cc.servicesPerHa) : 0;
    const d = isFinite(toNum(cc.otherPerHa)) ? Number(cc.otherPerHa) : 0;
    return a + b + c + d;
  }

  function renderTreatments() {
    const list = $("#treatmentsList");
    if (!list) return;
    list.innerHTML = "";

    const output = model.outputs[0]; // primary output for easy entry
    const outputId = output?.id;

    model.treatments.forEach(t => {
      const totalCostPerHa = computeTreatmentTotalCostPerHa(t);

      const card = document.createElement("div");
      card.className = "item";
      card.innerHTML = `
        <div class="item-head">
          <div class="item-title">${esc(t.name)}</div>
          <div class="item-actions">
            <button class="btn small ghost" data-act="del">Remove</button>
          </div>
        </div>

        <div class="row-4">
          <div class="field">
            <label data-tooltip="Treatment label used in results and exports.">Treatment name</label>
            <input type="text" value="${esc(t.name)}" data-k="name"/>
          </div>
          <div class="field">
            <label data-tooltip="Area (hectares) where the treatment is applied.">Area (ha)</label>
            <input type="number" step="0.01" value="${isFinite(toNum(t.area)) ? Number(t.area) : ""}" data-k="area"/>
          </div>
          <div class="field">
            <label data-tooltip="Treatment-specific adoption multiplier (0 to 1). Use this if not all area adopts.">Treatment adoption (0–1)</label>
            <input type="number" step="0.01" min="0" max="1" value="${isFinite(toNum(t.adoption)) ? Number(t.adoption) : 1}" data-k="adoption"/>
          </div>
          <div class="field">
            <label data-tooltip="Flag exactly one treatment as the control to compare all treatments against.">Control vs treatment</label>
            <select data-k="isControl">
              <option value="false" ${t.isControl ? "" : "selected"}>Treatment</option>
              <option value="true" ${t.isControl ? "selected" : ""}>Control</option>
            </select>
          </div>
        </div>

        <div class="row-4">
          <div class="field">
            <label data-tooltip="Capital cost in year 0 (one-off). This is applied at the start of the analysis.">Capital cost (AUD, year 0)</label>
            <input type="number" step="0.01" value="${isFinite(toNum(t.capitalCostYear0)) ? Number(t.capitalCostYear0) : ""}" data-k="capitalCostYear0"/>
          </div>
          <div class="field">
            <label data-tooltip="Treatment input cost only per hectare (as provided in the default dataset).">Input cost only ($/ha)</label>
            <input type="number" step="0.01" value="${isFinite(toNum(t.costComponents?.inputCostOnlyPerHa)) ? Number(t.costComponents.inputCostOnlyPerHa) : ""}" data-k="inputCostOnlyPerHa"/>
          </div>
          <div class="field">
            <label data-tooltip="Labour cost per hectare attributable to applying this treatment.">Labour ($/ha)</label>
            <input type="number" step="0.01" value="${isFinite(toNum(t.costComponents?.labourPerHa)) ? Number(t.costComponents.labourPerHa) : ""}" data-k="labourPerHa"/>
          </div>
          <div class="field">
            <label data-tooltip="Contracting or services cost per hectare attributable to this treatment.">Services ($/ha)</label>
            <input type="number" step="0.01" value="${isFinite(toNum(t.costComponents?.servicesPerHa)) ? Number(t.costComponents.servicesPerHa) : ""}" data-k="servicesPerHa"/>
          </div>
        </div>

        <div class="row-4">
          <div class="field">
            <label data-tooltip="Other per-hectare costs not covered above (for example transport, incidentals).">Other ($/ha)</label>
            <input type="number" step="0.01" value="${isFinite(toNum(t.costComponents?.otherPerHa)) ? Number(t.costComponents.otherPerHa) : ""}" data-k="otherPerHa"/>
          </div>
          <div class="field">
            <label data-tooltip="Total variable cost per hectare (computed from the components).">Total cost ($/ha)</label>
            <div class="metric">
              <div class="value">${money(totalCostPerHa)}</div>
              <div class="label">Computed automatically</div>
            </div>
          </div>
          <div class="field">
            <label data-tooltip="Number of years that treatment costs apply (starting year 0).">Cost duration (years)</label>
            <input type="number" step="1" min="0" value="${isFinite(toNum(t.costYears)) ? Number(t.costYears) : 1}" data-k="costYears"/>
          </div>
          <div class="field">
            <label data-tooltip="Number of years that benefits apply (starting year 1).">Benefit duration (years)</label>
            <input type="number" step="1" min="0" value="${isFinite(toNum(t.benefitYears)) ? Number(t.benefitYears) : model.time.years}" data-k="benefitYears"/>
          </div>
        </div>

        <div class="row-3">
          <div class="field">
            <label data-tooltip="Whether this cost is treated as constrained (useful for constrained-cost BCR in simulation).">Constrained cost</label>
            <select data-k="constrained">
              <option value="true" ${t.constrained ? "selected" : ""}>Yes</option>
              <option value="false" ${t.constrained ? "" : "selected"}>No</option>
            </select>
          </div>
          <div class="field">
            <label data-tooltip="Change in ${esc(output?.name || "output")} relative to control. Positive values increase benefits.">
              Delta ${esc(output?.name || "output")} (${esc(output?.unit || "unit")})
            </label>
            <input type="number" step="0.01" value="${isFinite(toNum(outputId ? t.deltas?.[outputId] : 0)) ? Number(t.deltas[outputId]) : 0}" data-k="deltaPrimary"/>
          </div>
          <div class="field">
            <label data-tooltip="Optional notes or assumptions specific to this treatment.">Notes</label>
            <input type="text" value="${esc(t.notes || "")}" data-k="notes"/>
          </div>
        </div>
      `;

      const delBtn = $('[data-act="del"]', card);
      delBtn.addEventListener("click", () => {
        const wasControl = t.isControl;
        model.treatments = model.treatments.filter(x => x.id !== t.id);
        if (wasControl) {
          // ensure we still have one control
          const fallback = model.treatments[0];
          if (fallback) fallback.isControl = true;
        }
        renderTreatments();
        refreshAll();
      });

      // Bind inputs
      $$("input[data-k], select[data-k]", card).forEach(inp => {
        inp.addEventListener("input", () => {
          const k = inp.dataset.k;

          if (k === "isControl") {
            const v = inp.value === "true";
            // enforce single control: set all false then set current true if selected
            if (v) model.treatments.forEach(x => (x.isControl = false));
            t.isControl = v;
            // If none selected, keep at least one control
            if (!model.treatments.some(x => x.isControl) && model.treatments[0]) model.treatments[0].isControl = true;
            renderTreatments();
            refreshAll();
            return;
          }

          if (k === "constrained") {
            t.constrained = inp.value === "true";
            refreshAll();
            return;
          }

          if (k === "deltaPrimary") {
            if (outputId) t.deltas[outputId] = toNum(inp.value);
            refreshAll();
            return;
          }

          if (k === "area") t.area = toNum(inp.value);
          else if (k === "adoption") t.adoption = clamp(toNum(inp.value), 0, 1);
          else if (k === "capitalCostYear0") t.capitalCostYear0 = toNum(inp.value);
          else if (k === "inputCostOnlyPerHa") t.costComponents.inputCostOnlyPerHa = toNum(inp.value);
          else if (k === "labourPerHa") t.costComponents.labourPerHa = toNum(inp.value);
          else if (k === "servicesPerHa") t.costComponents.servicesPerHa = toNum(inp.value);
          else if (k === "otherPerHa") t.costComponents.otherPerHa = toNum(inp.value);
          else if (k === "costYears") t.costYears = clamp(Math.round(toNum(inp.value) || 0), 0, 100);
          else if (k === "benefitYears") t.benefitYears = clamp(Math.round(toNum(inp.value) || 0), 0, 100);
          else if (k === "name") t.name = inp.value;
          else if (k === "notes") t.notes = inp.value;

          // Re-render to update computed total cost display cleanly
          renderTreatments();
          refreshAll();
        });
      });

      list.appendChild(card);
    });
  }

  function bindTreatmentsButtons() {
    const addBtn = $("#addTreatment");
    if (!addBtn) return;
    addBtn.addEventListener("click", () => {
      const outputId = model.outputs[0]?.id;
      model.treatments.push({
        id: uid(),
        name: "New treatment",
        area: 100,
        adoption: 1,
        isControl: false,
        constrained: true,
        notes: "",
        source: "User input",
        deltas: outputId ? { [outputId]: 0 } : {},
        capitalCostYear0: 0,
        costComponents: { inputCostOnlyPerHa: 0, labourPerHa: 0, servicesPerHa: 0, otherPerHa: 0 },
        costYears: 1,
        benefitYears: model.time.years
      });
      renderTreatments();
      refreshAll();
    });
  }

  // =========================
  // 11) RESULTS: SNAPSHOT + CONTROL COMPARISON TABLE
  // =========================
  function colourClassForBCR(bcr) {
    if (!isFinite(bcr)) return "";
    if (bcr >= 1) return "pos";
    return "neg";
  }

  function renderComparisonTable() {
    const summaryHost = $("#treatmentSummary");
    if (!summaryHost) return;

    const computed = computeAll();
    const table = computed.compareTable;
    const ranked = computed.perTreatment;
    const control = ranked.find(x => x.isControl) || null;

    // Snapshot headline: top performers and underperformers (no long narrative)
    const top3 = ranked.slice(0, 3);
    const bottom3 = ranked.slice(-3).reverse();

    const headline = `
      <div class="card subtle" style="margin-bottom:12px;">
        <h4 style="margin-bottom:8px;">Snapshot: economic performance (base case)</h4>
        <div class="row-2">
          <div class="field">
            <div class="small muted">Top treatments (by BCR)</div>
            <div class="small">
              ${top3
                .map(
                  x =>
                    `<span class="badge ${colourClassForBCR(x.bcr)}" title="BCR">${esc(x.name)}: ${isFinite(x.bcr) ? fmt(x.bcr, 3) : "n/a"}</span>`
                )
                .join(" ")}
            </div>
          </div>
          <div class="field">
            <div class="small muted">Lower-performing treatments (by BCR)</div>
            <div class="small">
              ${bottom3
                .map(
                  x =>
                    `<span class="badge ${colourClassForBCR(x.bcr)}" title="BCR">${esc(x.name)}: ${isFinite(x.bcr) ? fmt(x.bcr, 3) : "n/a"}</span>`
                )
                .join(" ")}
            </div>
          </div>
        </div>
        ${
          control
            ? `<div class="small muted" style="margin-top:10px;">All treatments are shown alongside the control to support transparent comparison. No treatments are hidden or excluded.</div>`
            : `<div class="small muted" style="margin-top:10px;">No control is currently flagged. Select one treatment as control in the Treatments tab to enable full control vs treatment comparison.</div>`
        }
      </div>
    `;

    // Build HTML table
    const colHeaders = table.cols
      .map(c => {
        const tag = c.isControl ? `<div class="small muted">Control</div>` : `<div class="small muted">Treatment</div>`;
        return `<th class="${c.isControl ? "is-control" : ""}">
          <div>${esc(c.name)}</div>
          ${tag}
        </th>`;
      })
      .join("");

    const bodyRows = table.rows
      .map((r, ri) => {
        const tip = tooltipForIndicator(r.key, r.label);
        const rowLabel = `<th class="row-label" data-tooltip="${esc(tip)}">${esc(r.label)}</th>`;

        const cells = table.cols
          .map((c, ci) => {
            const tr = ranked.find(x => x.id === c.id);
            const v = table.values[ri][ci];

            // Colour cues for key headline rows
            let cls = "";
            if (r.key === "bcr") cls = colourClassForBCR(v);
            if (r.key === "npv" || r.key === "inc_npv") cls = isFinite(v) ? (v >= 0 ? "pos" : "neg") : "";
            if (r.key === "rank") cls = c.isControl ? "is-control" : "";

            // Make control column visually distinct
            if (c.isControl) cls = (cls ? cls + " " : "") + "is-control";

            // Ensure copy/paste-friendly plain text content within cell
            const display = r.fmt(v);

            // Provide small subtext for delta rows
            const sub =
              r.key === "bcr" && tr
                ? `<div class="small muted">NPV: ${money(tr.npv)}</div>`
                : r.key === "rank" && tr
                ? `<div class="small muted">BCR: ${isFinite(tr.bcr) ? fmt(tr.bcr, 3) : "n/a"}</div>`
                : "";

            return `<td class="${esc(cls)}">${esc(display)}${sub}</td>`;
          })
          .join("");

        return `<tr>${rowLabel}${cells}</tr>`;
      })
      .join("");

    const tableHTML = `
      ${headline}
      <div class="table-scroll">
        <table class="summary-table" id="resultsCompareTable">
          <thead>
            <tr>
              <th class="row-label">Indicator</th>
              ${colHeaders}
            </tr>
          </thead>
          <tbody>${bodyRows}</tbody>
        </table>
      </div>

      <div class="small muted" style="margin-top:10px;">
        Tip: to copy into Word, select the table and copy. For Excel, use “Export Excel” in the Results tab footer.
      </div>
    `;

    summaryHost.innerHTML = tableHTML;
  }

  function tooltipForIndicator(key, label) {
    const defs = {
      "Present value of benefits":
        "Sum of discounted benefits over the analysis period using the base discount rate.",
      "Present value of costs":
        "Sum of discounted costs over the analysis period using the base discount rate.",
      "Net present value":
        "Present value of benefits minus present value of costs. Positive values indicate benefits exceed costs in present value terms.",
      "Benefit–cost ratio":
        "Present value of benefits divided by present value of costs. Values above 1 indicate benefits exceed costs (in PV terms).",
      "Return on investment":
        "Net present value as a percentage of present value of costs. Useful for comparing relative performance when scales differ.",
      "Ranking (BCR-based)":
        "Ordering from highest to lowest benefit–cost ratio under the current base case assumptions."
    };

    if (defs[label]) return defs[label];
    if (key === "inc_npv") return "Treatment NPV minus control NPV under the current assumptions.";
    if (key === "inc_pvBenefits") return "Treatment PV benefits minus control PV benefits under the current assumptions.";
    if (key === "inc_pvCosts") return "Treatment PV costs minus control PV costs under the current assumptions.";
    return "Indicator definition.";
  }

  function renderToplineMetrics() {
    // Keep existing tiles populated for compatibility, but the main comparison is the table.
    const computed = computeAll();
    const totals = computed.totals;

    const pvB = $("#pvBenefits");
    const pvC = $("#pvCosts");
    const npv = $("#npv");
    const bcr = $("#bcr");
    const roi = $("#roi");

    if (pvB) pvB.textContent = money(totals.pvBenefits);
    if (pvC) pvC.textContent = money(totals.pvCosts);
    if (npv) npv.textContent = money(totals.npv);
    if (bcr) bcr.textContent = isFinite(totals.bcr) ? fmt(totals.bcr, 3) : "n/a";
    if (roi) roi.textContent = isFinite(totals.roi) ? percent(totals.roi) : "n/a";

    // Optional: fill control and combined group blocks using perTreatment
    const control = computed.perTreatment.find(x => x.isControl) || null;
    const nonControl = computed.perTreatment.filter(x => !x.isControl);

    const setIf = (id, v, formatter) => {
      const el = $("#" + id);
      if (el) el.textContent = formatter(v);
    };

    if (control) {
      setIf("pvBenefitsControl", control.pvBenefits, money);
      setIf("pvCostsControl", control.pvCosts, money);
      setIf("npvControl", control.npv, money);
      setIf("bcrControl", control.bcr, v => (isFinite(v) ? fmt(v, 3) : "n/a"));
      setIf("roiControl", control.roi, v => (isFinite(v) ? percent(v) : "n/a"));
    } else {
      ["pvBenefitsControl", "pvCostsControl", "npvControl", "bcrControl", "roiControl"].forEach(id =>
        setIf(id, NaN, () => "n/a")
      );
    }

    if (nonControl.length) {
      const pvBt = nonControl.reduce((s, x) => s + (isFinite(x.pvBenefits) ? x.pvBenefits : 0), 0);
      const pvCt = nonControl.reduce((s, x) => s + (isFinite(x.pvCosts) ? x.pvCosts : 0), 0);
      const npvt = pvBt - pvCt;
      const bcrt = pvCt !== 0 ? pvBt / pvCt : NaN;
      const roit = pvCt !== 0 ? (npvt / pvCt) * 100 : NaN;

      setIf("pvBenefitsTreat", pvBt, money);
      setIf("pvCostsTreat", pvCt, money);
      setIf("npvTreat", npvt, money);
      setIf("bcrTreat", bcrt, v => (isFinite(v) ? fmt(v, 3) : "n/a"));
      setIf("roiTreat", roit, v => (isFinite(v) ? percent(v) : "n/a"));
    } else {
      ["pvBenefitsTreat", "pvCostsTreat", "npvTreat", "bcrTreat", "roiTreat"].forEach(id =>
        setIf(id, NaN, () => "n/a")
      );
    }
  }

  function renderTimeProjection() {
    const tbody = $("#timeProjectionTable tbody");
    if (!tbody) return;
    const computed = computeAll();
    const rows = computed.timeProjection;

    tbody.innerHTML = rows
      .map(
        r => `
      <tr>
        <td>${esc(String(r.years))}</td>
        <td>${esc(money(r.pvBenefits))}</td>
        <td>${esc(money(r.pvCosts))}</td>
        <td>${esc(money(r.npv))}</td>
        <td>${esc(isFinite(r.bcr) ? fmt(r.bcr, 3) : "n/a")}</td>
      </tr>
    `
      )
      .join("");
  }

  // =========================
  // 12) EXPORTS: EXCEL + CSV + PDF
  // =========================
  function buildResultsMatrixForExport() {
    const computed = computeAll();
    const table = computed.compareTable;

    // 2D array with header row
    const header = ["Indicator"].concat(table.cols.map(c => c.name));
    const aoa = [header];

    table.rows.forEach((r, ri) => {
      const row = [r.label];
      table.cols.forEach((c, ci) => {
        const v = table.values[ri][ci];
        // export raw numeric where possible, else text
        row.push(isFinite(v) ? v : null);
      });
      aoa.push(row);
    });

    return { aoa, computed };
  }

  function exportExcel() {
    if (!ensureXLSX()) return;

    const { aoa, computed } = buildResultsMatrixForExport();

    const wb = XLSX.utils.book_new();

    // Results sheet
    const wsRes = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, wsRes, "Results_Compare");

    // Inputs: Treatments
    const tHeader = [
      "Treatment name",
      "Is control",
      "Area (ha)",
      "Treatment adoption",
      "Capital cost (year 0, AUD)",
      "Input cost only ($/ha)",
      "Labour ($/ha)",
      "Services ($/ha)",
      "Other ($/ha)",
      "Total variable cost ($/ha)",
      "Cost duration (years)",
      "Benefit duration (years)",
      "Delta Grain yield (t/ha)",
      "Notes",
      "Source"
    ];
    const outputId = model.outputs[0]?.id;
    const tRows = model.treatments.map(t => [
      t.name,
      t.isControl ? 1 : 0,
      t.area,
      t.adoption,
      t.capitalCostYear0,
      t.costComponents?.inputCostOnlyPerHa,
      t.costComponents?.labourPerHa,
      t.costComponents?.servicesPerHa,
      t.costComponents?.otherPerHa,
      computeTreatmentTotalCostPerHa(t),
      t.costYears,
      t.benefitYears,
      outputId ? t.deltas?.[outputId] : 0,
      t.notes || "",
      t.source || ""
    ]);
    const wsTreat = XLSX.utils.aoa_to_sheet([tHeader, ...tRows]);
    XLSX.utils.book_append_sheet(wb, wsTreat, "Treatments");

    // Outputs sheet
    const oHeader = ["Output name", "Unit", "Value per unit (AUD)", "Source"];
    const oRows = model.outputs.map(o => [o.name, o.unit, o.value, o.source || ""]);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([oHeader, ...oRows]), "Outputs");

    // Settings sheet
    const wsSet = XLSX.utils.aoa_to_sheet([
      ["Tool name", TOOL_NAME],
      ["Version", TOOL_VERSION],
      ["Project name", model.project.name],
      ["Analysis years", model.time.years],
      ["Discount rate (base, %)", model.time.discBase],
      ["Adoption (base)", model.adoption.base],
      ["Risk (base)", model.risk.base],
      ["Last run", computed.lastRun || ""]
    ]);
    XLSX.utils.book_append_sheet(wb, wsSet, "Scenario");

    // Raw data sheet (full, unmodified)
    const rawLines = model.raw.fabaBeans2022.split(/\r?\n/);
    const rawAOA = rawLines.map(line => line.split("\t"));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rawAOA), "RawData_FabaBeans_2022");

    const filename = `${slug(TOOL_NAME)}_${slug(model.project.name)}_results.xlsx`;
    XLSX.writeFile(wb, filename);
  }

  function exportCsv() {
    // CSV export of comparison table (copy/paste into Word still easiest directly from UI)
    const { aoa } = buildResultsMatrixForExport();
    const csv = aoa
      .map(row =>
        row
          .map(v => {
            if (v === null || v === undefined) return "";
            const s = String(v);
            if (s.includes(",") || s.includes('"') || s.includes("\n")) return `"${s.replace(/"/g, '""')}"`;
            return s;
          })
          .join(",")
      )
      .join("\n");

    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${slug(TOOL_NAME)}_${slug(model.project.name)}_results.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  function bindExportButtons() {
    const exportCsvBtn = $("#exportCsv");
    const exportCsvFootBtn = $("#exportCsvFoot");

    if (exportCsvBtn) {
      exportCsvBtn.textContent = "Export Excel";
      exportCsvBtn.addEventListener("click", exportExcel);
    }
    if (exportCsvFootBtn) {
      exportCsvFootBtn.textContent = "Export Excel";
      exportCsvFootBtn.addEventListener("click", exportExcel);
    }

    const exportPdfBtn = $("#exportPdf");
    const exportPdfFootBtn = $("#exportPdfFoot");

    const doPrint = () => window.print();
    if (exportPdfBtn) exportPdfBtn.addEventListener("click", doPrint);
    if (exportPdfFootBtn) exportPdfFootBtn.addEventListener("click", doPrint);
  }

  // =========================
  // 13) EXCEL-FIRST WORKFLOW (template + import)
  // =========================
  function buildTemplateWorkbook({ includeSample }) {
    if (!ensureXLSX()) return null;

    const wb = XLSX.utils.book_new();

    // ReadMe sheet (plain language)
    const readme = [
      [TOOL_NAME, ""],
      ["Version", TOOL_VERSION],
      ["", ""],
      ["How to use this Excel workflow", ""],
      [
        "1) Download this workbook. Edit values in the Outputs and Treatments sheets (and optionally Scenario).",
        ""
      ],
      ["2) Save the file. Return to the tool and import it in the Excel tab.", ""],
      ["3) The tool will validate structure and apply inputs automatically.", ""],
      ["", ""],
      ["Key expectations", ""],
      ["- Keep sheet names and column headers unchanged.", ""],
      ["- Use numbers (no $ signs needed). Missing values can be left blank.", ""],
      ["- Exactly one treatment should have Is control = 1.", ""],
      ["", ""],
      ["Notes", ""],
      ["- The RawData sheet contains the full default dataset for reference and auditing.", ""]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(readme), "ReadMe");

    // Scenario sheet
    const scenario = [
      ["Project name", model.project.name],
      ["Analysis years", model.time.years],
      ["Discount rate (base, %)", model.time.discBase],
      ["Adoption (base, 0-1)", model.adoption.base],
      ["Risk (base, 0-1)", model.risk.base]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(scenario), "Scenario");

    // Outputs sheet
    const oHeader = ["Output name", "Unit", "Value per unit (AUD)", "Source"];
    const oRows = (includeSample ? model.outputs : [{ name: "Grain yield", unit: "t/ha", value: 0, source: "" }]).map(o => [
      o.name,
      o.unit,
      o.value,
      o.source || ""
    ]);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([oHeader, ...oRows]), "Outputs");

    // Treatments sheet
    const tHeader = [
      "Treatment name",
      "Is control",
      "Area (ha)",
      "Treatment adoption",
      "Capital cost (year 0, AUD)",
      "Input cost only ($/ha)",
      "Labour ($/ha)",
      "Services ($/ha)",
      "Other ($/ha)",
      "Cost duration (years)",
      "Benefit duration (years)",
      "Delta Grain yield (t/ha)",
      "Notes",
      "Source"
    ];
    const outputId = model.outputs[0]?.id;

    const treatmentsForSheet = includeSample
      ? model.treatments
      : [
          {
            name: "Control",
            isControl: 1,
            area: 100,
            adoption: 1,
            capitalCostYear0: 0,
            inputCostOnlyPerHa: 0,
            labourPerHa: 0,
            servicesPerHa: 0,
            otherPerHa: 0,
            costYears: 1,
            benefitYears: model.time.years,
            deltaYield: 0,
            notes: "",
            source: ""
          }
        ];

    const tRows = treatmentsForSheet.map(t => [
      t.name,
      t.isControl ? 1 : 0,
      isFinite(toNum(t.area)) ? Number(t.area) : 0,
      isFinite(toNum(t.adoption)) ? Number(t.adoption) : 1,
      isFinite(toNum(t.capitalCostYear0)) ? Number(t.capitalCostYear0) : 0,
      isFinite(toNum(t.costComponents?.inputCostOnlyPerHa ?? t.inputCostOnlyPerHa)) ? Number(t.costComponents?.inputCostOnlyPerHa ?? t.inputCostOnlyPerHa) : 0,
      isFinite(toNum(t.costComponents?.labourPerHa ?? t.labourPerHa)) ? Number(t.costComponents?.labourPerHa ?? t.labourPerHa) : 0,
      isFinite(toNum(t.costComponents?.servicesPerHa ?? t.servicesPerHa)) ? Number(t.costComponents?.servicesPerHa ?? t.servicesPerHa) : 0,
      isFinite(toNum(t.costComponents?.otherPerHa ?? t.otherPerHa)) ? Number(t.costComponents?.otherPerHa ?? t.otherPerHa) : 0,
      isFinite(toNum(t.costYears)) ? Number(t.costYears) : 1,
      isFinite(toNum(t.benefitYears)) ? Number(t.benefitYears) : model.time.years,
      outputId ? (isFinite(toNum(t.deltas?.[outputId])) ? Number(t.deltas[outputId]) : 0) : 0,
      t.notes || "",
      t.source || ""
    ]);

    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([tHeader, ...tRows]), "Treatments");

    // Raw data sheet (always included; full default dataset)
    const rawLines = model.raw.fabaBeans2022.split(/\r?\n/);
    const rawAOA = rawLines.map(line => line.split("\t"));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rawAOA), "RawData_FabaBeans_2022");

    return wb;
  }

  function downloadTemplate(blank) {
    if (!ensureXLSX()) return;
    const wb = buildTemplateWorkbook({ includeSample: !blank });
    if (!wb) return;
    const filename = blank
      ? `${slug(TOOL_NAME)}_excel_template_blank.xlsx`
      : `${slug(TOOL_NAME)}_excel_template_sample.xlsx`;
    XLSX.writeFile(wb, filename);
  }

  function getExcelFileInput() {
    let input = $("#excelFileInputHidden");
    if (!input) {
      input = document.createElement("input");
      input.type = "file";
      input.accept = ".xlsx,.xls";
      input.id = "excelFileInputHidden";
      input.style.display = "none";
      document.body.appendChild(input);
    }
    return input;
  }

  function parseExcelFile(file) {
    if (!ensureXLSX()) return Promise.resolve(null);

    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error("Failed to read file."));
      reader.onload = e => {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: "array" });

          // Expect at least Outputs and Treatments
          const hasOutputs = wb.Sheets["Outputs"];
          const hasTreatments = wb.Sheets["Treatments"];
          if (!hasOutputs || !hasTreatments) {
            throw new Error("Missing required sheets. Expected sheets named 'Outputs' and 'Treatments'.");
          }

          const outputs = XLSX.utils.sheet_to_json(hasOutputs, { defval: "" });
          const treatments = XLSX.utils.sheet_to_json(hasTreatments, { defval: "" });
          const scenario = wb.Sheets["Scenario"] ? XLSX.utils.sheet_to_json(wb.Sheets["Scenario"], { header: 1 }) : null;

          resolve({ wb, outputs, treatments, scenario });
        } catch (err) {
          reject(err);
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  function validateParsedExcel(parsed) {
    const errors = [];

    // Outputs: need at least one row with Output name and Value
    if (!Array.isArray(parsed.outputs) || parsed.outputs.length < 1) {
      errors.push("Outputs sheet has no rows.");
    } else {
      const first = parsed.outputs[0];
      if (!("Output name" in first) || !("Value per unit (AUD)" in first)) {
        errors.push("Outputs sheet is missing required columns (Output name, Value per unit (AUD)).");
      }
    }

    // Treatments: need columns and at least one control
    if (!Array.isArray(parsed.treatments) || parsed.treatments.length < 1) {
      errors.push("Treatments sheet has no rows.");
    } else {
      const cols = parsed.treatments[0];
      const required = [
        "Treatment name",
        "Is control",
        "Area (ha)",
        "Treatment adoption",
        "Capital cost (year 0, AUD)",
        "Input cost only ($/ha)",
        "Cost duration (years)",
        "Benefit duration (years)",
        "Delta Grain yield (t/ha)"
      ];
      required.forEach(c => {
        if (!(c in cols)) errors.push(`Treatments sheet is missing required column: ${c}`);
      });

      const nControl = parsed.treatments.reduce((s, r) => s + (toNum(r["Is control"]) === 1 ? 1 : 0), 0);
      if (nControl !== 1) errors.push("Treatments sheet must have exactly one control (Is control = 1).");
    }

    return errors;
  }

  function applyParsedExcel(parsed) {
    // Outputs
    const outputs = parsed.outputs
      .map(r => ({
        id: uid(),
        name: String(r["Output name"] || "").trim(),
        unit: String(r["Unit"] || "").trim(),
        value: toNum(r["Value per unit (AUD)"]),
        source: String(r["Source"] || "").trim()
      }))
      .filter(o => o.name.length > 0);

    if (outputs.length === 0) {
      showToast("No valid outputs found in Excel. Nothing applied.");
      return;
    }
    model.outputs = outputs;

    // Scenario (optional)
    if (Array.isArray(parsed.scenario) && parsed.scenario.length) {
      // Scenario is a two-column layout in our template
      const dict = {};
      parsed.scenario.forEach(row => {
        if (Array.isArray(row) && row.length >= 2) dict[String(row[0]).trim()] = row[1];
      });
      if (dict["Project name"]) model.project.name = String(dict["Project name"]);
      if (isFinite(toNum(dict["Analysis years"]))) model.time.years = toNum(dict["Analysis years"]);
      if (isFinite(toNum(dict["Discount rate (base, %)"]))) model.time.discBase = toNum(dict["Discount rate (base, %)"]);
      if (isFinite(toNum(dict["Adoption (base, 0-1)"]))) model.adoption.base = clamp(toNum(dict["Adoption (base, 0-1)"]), 0, 1);
      if (isFinite(toNum(dict["Risk (base, 0-1)"]))) model.risk.base = clamp(toNum(dict["Risk (base, 0-1)"]), 0, 1);
    }

    // Treatments
    const primaryOutputId = model.outputs[0].id;
    const treatments = parsed.treatments
      .map(r => {
        const name = String(r["Treatment name"] || "").trim();
        if (!name) return null;
        return {
          id: uid(),
          name,
          isControl: toNum(r["Is control"]) === 1,
          area: toNum(r["Area (ha)"]) || 0,
          adoption: clamp(toNum(r["Treatment adoption"]) || 1, 0, 1),
          capitalCostYear0: toNum(r["Capital cost (year 0, AUD)"]) || 0,
          costComponents: {
            inputCostOnlyPerHa: toNum(r["Input cost only ($/ha)"]),
            labourPerHa: toNum(r["Labour ($/ha)"]),
            servicesPerHa: toNum(r["Services ($/ha)"]),
            otherPerHa: toNum(r["Other ($/ha)"])
          },
          costYears: clamp(Math.round(toNum(r["Cost duration (years)"]) || 0), 0, 100),
          benefitYears: clamp(Math.round(toNum(r["Benefit duration (years)"]) || 0), 0, 100),
          deltas: { [primaryOutputId]: toNum(r["Delta Grain yield (t/ha)"]) || 0 },
          notes: String(r["Notes"] || ""),
          source: String(r["Source"] || "Excel import"),
          constrained: true
        };
      })
      .filter(Boolean);

    model.treatments = treatments;

    // Ensure exactly one control (already validated)
    initTreatmentDeltas();

    // Refresh UI
    bindProjectFields();
    bindSettingsFields();
    renderOutputs();
    renderTreatments();
    refreshAll();
    showToast("Excel data applied successfully.");
  }

  function bindExcelButtons() {
    const downloadTemplateBtn = $("#downloadTemplate");
    const downloadSampleBtn = $("#downloadSample");
    const parseBtn = $("#parseExcel");
    const importBtn = $("#importExcel");

    if (downloadTemplateBtn) {
      downloadTemplateBtn.addEventListener("click", () => downloadTemplate(true));
    }
    if (downloadSampleBtn) {
      downloadSampleBtn.addEventListener("click", () => downloadTemplate(false));
    }

    const pickFile = () => {
      const input = getExcelFileInput();
      input.value = "";
      input.click();
      return new Promise(resolve => {
        input.onchange = () => resolve(input.files && input.files[0] ? input.files[0] : null);
      });
    };

    if (parseBtn) {
      parseBtn.addEventListener("click", async () => {
        const file = await pickFile();
        if (!file) return;

        try {
          const parsed = await parseExcelFile(file);
          const errors = validateParsedExcel(parsed);
          if (errors.length) {
            parsedExcel = null;
            showToast("Excel parse failed. See details in console.");
            console.error("Excel validation errors:", errors);
            alert("Excel import validation failed:\n\n" + errors.join("\n"));
            return;
          }

          parsedExcel = parsed;
          showToast("Excel parsed successfully. Click ‘Apply parsed Excel data’ to update the tool.");

        } catch (err) {
          parsedExcel = null;
          console.error(err);
          alert("Excel parse failed: " + (err && err.message ? err.message : String(err)));
        }
      });
    }

    if (importBtn) {
      importBtn.addEventListener("click", () => {
        if (!parsedExcel) {
          showToast("No parsed Excel data found. Click ‘Parse Excel file’ first.");
          return;
        }
        applyParsedExcel(parsedExcel);
        parsedExcel = null;
      });
    }
  }

  // =========================
  // 14) AI INTERPRETATION HELPER (NON-PRESCRIPTIVE + LEARNING AID)
  // =========================
  function buildAIInterpretationPrompt() {
    const computed = computeAll();
    const table = computed.compareTable;
    const ranked = computed.perTreatment;
    const control = ranked.find(x => x.isControl) || null;

    const output = model.outputs[0];
    const price = isFinite(toNum(output?.value)) ? Number(output.value) : NaN;

    // Build a compact, model-agnostic prompt that can be pasted anywhere.
    // No dictated decisions; explain drivers; suggest realistic improvement levers for low BCR.
    const treatmentBlocks = ranked
      .map(x => {
        const t = x.t;
        const varCostPerHa = computeTreatmentTotalCostPerHa(t);
        const deltaYield = isFinite(toNum(t.deltas?.[output?.id])) ? Number(t.deltas[output.id]) : 0;

        const drivers = [
          `Delta ${output?.name || "output"}: ${isFinite(deltaYield) ? fmt(deltaYield, 3) : "n/a"} ${output?.unit || ""}`,
          `Value per unit used: ${isFinite(price) ? money(price) : "n/a"}`,
          `Area: ${isFinite(toNum(t.area)) ? fmt(Number(t.area), 2) : "n/a"} ha`,
          `Treatment adoption: ${isFinite(toNum(t.adoption)) ? fmt(Number(t.adoption), 2) : "n/a"}`,
          `Variable cost: ${money(varCostPerHa)} per ha`,
          `Capital cost (year 0): ${money(isFinite(toNum(t.capitalCostYear0)) ? Number(t.capitalCostYear0) : NaN)}`,
          `Cost duration: ${isFinite(toNum(t.costYears)) ? String(t.costYears) : "n/a"} years`,
          `Benefit duration: ${isFinite(toNum(t.benefitYears)) ? String(t.benefitYears) : "n/a"} years`
        ].join("; ");

        const perf = `PV benefits ${money(x.pvBenefits)}, PV costs ${money(x.pvCosts)}, NPV ${money(
          x.npv
        )}, BCR ${isFinite(x.bcr) ? fmt(x.bcr, 3) : "n/a"}, ROI ${isFinite(x.roi) ? fmt(x.roi, 2) + "%" : "n/a"}, Rank ${
          x.rank
        }.`;

        // Improvement suggestions for low BCR (non-prescriptive)
        let improve = "";
        if (isFinite(x.bcr) && x.bcr < 1) {
          improve = [
            "Learning and improvement ideas (not rules):",
            "Consider whether costs are overstated or benefits understated for this treatment in your context.",
            "If cost is the main driver, explore lowering input costs (cheaper product, reduced application rate, combining passes, contracting options) or reducing labour/services through operational changes.",
            "If benefits are the main driver, explore options that plausibly increase yield response (better timing, improved establishment, improved nutrient matching, addressing limiting constraints, or targeting the treatment to soil zones where response is higher).",
            "If price assumptions drive results, test a realistic price range or quality premiums that could apply.",
            "If persistence is uncertain, test shorter benefit duration or staged adoption to reflect learning over time.",
            "Use the simulation tab to test sensitivity rather than relying on a single set of assumptions."
          ].join("\n");
        } else if (isFinite(x.bcr) && x.bcr >= 1) {
          improve = [
            "Learning and improvement ideas (not rules):",
            "Stress-test assumptions that drive performance (yield response, price/value, cost components, benefit duration, adoption, and risk).",
            "Check operational feasibility and whether scaling effects might change costs or benefits.",
            "Use simulation and scenario variants to identify the conditions under which this treatment remains attractive."
          ].join("\n");
        }

        return [
          `Treatment: ${t.name}${t.isControl ? " (CONTROL)" : ""}`,
          `Performance: ${perf}`,
          `Drivers: ${drivers}`,
          improve
        ].join("\n");
      })
      .join("\n\n---\n\n");

    // Include the comparison table values as plain text for the LLM
    const colNames = table.cols.map(c => c.name);
    const tableText = [
      ["Indicator", ...colNames].join("\t"),
      ...table.rows.map((r, ri) => {
        const row = [r.label];
        for (let ci = 0; ci < colNames.length; ci++) row.push(isFinite(table.values[ri][ci]) ? String(table.values[ri][ci]) : "");
        return row.join("\t");
      })
    ].join("\n");

    const prompt = [
      `You are assisting with interpretation of cost–benefit analysis (CBA) results produced by ${TOOL_NAME}.`,
      `Write a plain-English interpretation suitable for farmers, agronomists, and decision-makers. Do not dictate decisions or apply arbitrary thresholds. Explain what the indicators mean, what drives differences, and what trade-offs may matter.`,
      ``,
      `Context and settings:`,
      `- Project: ${model.project.name}`,
      `- Analysis years: ${model.time.years}`,
      `- Discount rate (base): ${model.time.discBase}%`,
      `- Adoption (base): ${fmt(model.adoption.base, 2)}`,
      `- Risk (base): ${fmt(model.risk.base, 2)} (benefits are reduced by (1 - risk))`,
      `- Primary output valued: ${output?.name || "Output"} at ${isFinite(price) ? money(price) : "n/a"} per ${output?.unit || "unit"}`,
      control ? `- Control treatment: ${control.name}` : `- Control treatment: Not specified (interpret comparisons cautiously).`,
      ``,
      `Required structure:`,
      `1) Brief overview of what the table shows and how to read it (control vs treatments).`,
      `2) Explain each headline indicator (PV benefits, PV costs, NPV, BCR, ROI, ranking) in simple terms.`,
      `3) Identify which treatments perform better or worse and explain why, using the drivers provided.`,
      `4) Highlight trade-offs (e.g., higher costs but higher benefit, risk exposure, scale effects).`,
      `5) For any low BCR treatments, suggest realistic, practical improvement levers (cost reduction, yield response improvements, price/quality changes, agronomic targeting, timing, operational efficiency). Frame as reflection and learning, not rules.`,
      `6) End with a short section listing assumptions to stress-test (yield response, price, costs, benefit duration, adoption, risk, discount rate).`,
      ``,
      `Comparison table (tab-separated, numeric values where available):`,
      tableText,
      ``,
      `Treatment-by-treatment detail (drivers + performance):`,
      treatmentBlocks
    ].join("\n");

    return prompt;
  }

  function refreshCopilotPreview() {
    const preview = $("#copilotPreview");
    if (!preview) return;
    preview.value = buildAIInterpretationPrompt();
  }

  function bindCopilot() {
    const btn = $("#openCopilot");
    const tabTitle = $("#tab-copilot h2");
    if (tabTitle) tabTitle.textContent = "AI interpretation helper (copy/paste prompt)";

    if (!btn) return;
    btn.textContent = "Copy AI interpretation prompt";
    btn.addEventListener("click", async () => {
      const text = buildAIInterpretationPrompt();
      try {
        await navigator.clipboard.writeText(text);
        showToast("AI interpretation prompt copied. Paste into Copilot, ChatGPT, or your preferred assistant.");
      } catch {
        // fallback
        const preview = $("#copilotPreview");
        if (preview) {
          preview.removeAttribute("readonly");
          preview.select();
          document.execCommand("copy");
          preview.setAttribute("readonly", "readonly");
          showToast("AI interpretation prompt copied.");
        } else {
          showToast("Copy failed. Please select the text in the preview box and copy manually.");
        }
      }
      refreshCopilotPreview();
    });

    refreshCopilotPreview();
  }

  // =========================
  // 15) SIMULATION (kept compatible; uses current model)
  // =========================
  function runSimulation() {
    const simN = clamp(Math.round(toNum($("#simN")?.value) || model.sim.n), 100, 200000);
    model.sim.n = simN;

    const target = toNum($("#targetBCR")?.value);
    if (isFinite(target)) model.sim.targetBCR = target;

    const bcrModeEl = $("#bcrMode");
    if (bcrModeEl) model.sim.bcrMode = bcrModeEl.value || "all";

    const seed = toNum($("#randSeed")?.value);
    model.sim.seed = isFinite(seed) ? seed : null;

    const varPct = toNum($("#simVarPct")?.value);
    model.sim.variationPct = isFinite(varPct) ? clamp(varPct, 0, 100) : model.sim.variationPct;

    const varyOutputs = $("#simVaryOutputs");
    const varyTreatCosts = $("#simVaryTreatCosts");
    const varyInputCosts = $("#simVaryInputCosts");
    if (varyOutputs) model.sim.varyOutputs = varyOutputs.value === "true";
    if (varyTreatCosts) model.sim.varyTreatCosts = varyTreatCosts.value === "true";
    if (varyInputCosts) model.sim.varyInputCosts = varyInputCosts.value === "true";

    const rand = rng(model.sim.seed || undefined);

    const baseYears = clamp(Math.round(toNum(model.time.years) || 10), 1, 100);
    const discLow = toNum(model.time.discLow) || 4;
    const discBase = toNum(model.time.discBase) || 7;
    const discHigh = toNum(model.time.discHigh) || 10;

    const riskLow = clamp(toNum(model.risk.low) || 0, 0, 1);
    const riskBase = clamp(toNum(model.risk.base) || 0, 0, 1);
    const riskHigh = clamp(toNum(model.risk.high) || 0, 0, 1);

    const adoptLow = clamp(toNum(model.adoption.low) || 0, 0, 1);
    const adoptBase = clamp(toNum(model.adoption.base) || 0, 0, 1);
    const adoptHigh = clamp(toNum(model.adoption.high) || 1, 0, 1);

    const varMult = model.sim.variationPct / 100;

    const baseOutputVal = isFinite(toNum(model.outputs[0]?.value)) ? Number(model.outputs[0].value) : 0;

    const npvArr = [];
    const bcrArr = [];

    const status = $("#simStatus");
    if (status) status.textContent = "Running simulation...";

    for (let i = 0; i < simN; i++) {
      const disc = triangular(rand(), discLow, discBase, discHigh);
      const risk = triangular(rand(), riskLow, riskBase, riskHigh);
      const adopt = triangular(rand(), adoptLow, adoptBase, adoptHigh);

      // Perturb output value
      const outVal = model.sim.varyOutputs ? baseOutputVal * (1 + (rand() * 2 - 1) * varMult) : baseOutputVal;

      // Temporarily apply output value
      const saved = model.outputs[0].value;
      model.outputs[0].value = outVal;

      // Perturb treatment costs if requested
      const savedCosts = model.treatments.map(t => ({
        id: t.id,
        input: t.costComponents?.inputCostOnlyPerHa,
        labour: t.costComponents?.labourPerHa,
        services: t.costComponents?.servicesPerHa,
        other: t.costComponents?.otherPerHa
      }));

      if (model.sim.varyTreatCosts) {
        model.treatments.forEach(t => {
          const cc = t.costComponents || {};
          const bump = x => (isFinite(toNum(x)) ? Number(x) * (1 + (rand() * 2 - 1) * varMult) : x);
          cc.inputCostOnlyPerHa = bump(cc.inputCostOnlyPerHa);
          cc.labourPerHa = bump(cc.labourPerHa);
          cc.servicesPerHa = bump(cc.servicesPerHa);
          cc.otherPerHa = bump(cc.otherPerHa);
          t.costComponents = cc;
        });
      }

      // Compute outcomes
      const per = model.treatments.map(t => calcPerTreatmentCashflows(t, { years: baseYears, discBase: disc, riskBase: risk, adoptBase: adopt }));
      const totals = calcProjectTotals(per);

      npvArr.push(totals.npv);
      bcrArr.push(totals.bcr);

      // Restore
      model.outputs[0].value = saved;
      if (model.sim.varyTreatCosts) {
        model.treatments.forEach(t => {
          const s = savedCosts.find(z => z.id === t.id);
          if (!s) return;
          t.costComponents.inputCostOnlyPerHa = s.input;
          t.costComponents.labourPerHa = s.labour;
          t.costComponents.servicesPerHa = s.services;
          t.costComponents.otherPerHa = s.other;
        });
      }
    }

    model.sim.results = { npv: npvArr, bcr: bcrArr };

    renderSimulationSummary();

    if (status) status.textContent = `Simulation complete (${simN.toLocaleString()} runs).`;
    showToast("Simulation complete.");
  }

  function quantile(arr, q) {
    if (!arr.length) return NaN;
    const s = arr.slice().sort((a, b) => a - b);
    const pos = (s.length - 1) * q;
    const base = Math.floor(pos);
    const rest = pos - base;
    if (s[base + 1] === undefined) return s[base];
    return s[base] + rest * (s[base + 1] - s[base]);
  }

  function renderSimulationSummary() {
    const npv = model.sim.results.npv || [];
    const bcr = model.sim.results.bcr || [];

    const set = (id, v, f = money) => {
      const el = $("#" + id);
      if (el) el.textContent = f(v);
    };

    set("simNpvMin", Math.min(...npv));
    set("simNpvMax", Math.max(...npv));
    set("simNpvMean", npv.reduce((s, x) => s + x, 0) / (npv.length || 1));
    set("simNpvMedian", quantile(npv, 0.5));
    set("simNpvProb", npv.filter(x => x > 0).length / (npv.length || 1), v => (isFinite(v) ? fmt(100 * v, 1) + "%" : "n/a"));

    set("simBcrMin", Math.min(...bcr), v => (isFinite(v) ? fmt(v, 3) : "n/a"));
    set("simBcrMax", Math.max(...bcr), v => (isFinite(v) ? fmt(v, 3) : "n/a"));
    set("simBcrMean", bcr.reduce((s, x) => s + x, 0) / (bcr.length || 1), v => (isFinite(v) ? fmt(v, 3) : "n/a"));
    set("simBcrMedian", quantile(bcr, 0.5), v => (isFinite(v) ? fmt(v, 3) : "n/a"));

    const prob1 = bcr.filter(x => x > 1).length / (bcr.length || 1);
    const probT = bcr.filter(x => x > model.sim.targetBCR).length / (bcr.length || 1);

    set("simBcrProb1", prob1, v => (isFinite(v) ? fmt(100 * v, 1) + "%" : "n/a"));
    set("simBcrProbTarget", probT, v => (isFinite(v) ? fmt(100 * v, 1) + "%" : "n/a"));

    const tgtLabel = $("#simBcrTargetLabel");
    if (tgtLabel) tgtLabel.textContent = String(model.sim.targetBCR);

    // Charts: keep existing canvases if present; if not, ignore.
    // We avoid styling overrides; use basic bins.
    drawHistogram("histNpv", npv);
    drawHistogram("histBcr", bcr);
  }

  function drawHistogram(canvasId, data) {
    const canvas = $("#" + canvasId);
    if (!canvas || !canvas.getContext) return;
    const ctx = canvas.getContext("2d");
    const w = canvas.width;
    const h = canvas.height;
    ctx.clearRect(0, 0, w, h);

    if (!data || data.length < 10) {
      ctx.fillText("Not enough data to plot.", 10, 20);
      return;
    }

    const nBins = 25;
    const min = Math.min(...data);
    const max = Math.max(...data);
    const span = max - min || 1;
    const bins = new Array(nBins).fill(0);
    for (const x of data) {
      const idx = clamp(Math.floor(((x - min) / span) * nBins), 0, nBins - 1);
      bins[idx] += 1;
    }
    const maxBin = Math.max(...bins) || 1;

    // Axes
    const pad = 24;
    ctx.strokeRect(pad, pad, w - 2 * pad, h - 2 * pad);

    // Bars
    const barW = (w - 2 * pad) / nBins;
    for (let i = 0; i < nBins; i++) {
      const bh = ((h - 2 * pad) * bins[i]) / maxBin;
      ctx.fillRect(pad + i * barW, h - pad - bh, Math.max(1, barW - 1), bh);
    }
  }

  function bindSimulation() {
    const runBtn = $("#runSim");
    if (runBtn) runBtn.addEventListener("click", runSimulation);

    const simN = $("#simN");
    if (simN) simN.value = model.sim.n;

    const target = $("#targetBCR");
    if (target) target.value = model.sim.targetBCR;

    const bcrMode = $("#bcrMode");
    if (bcrMode) bcrMode.value = model.sim.bcrMode;

    const seed = $("#randSeed");
    if (seed) seed.value = model.sim.seed ?? "";

    const varPct = $("#simVarPct");
    if (varPct) varPct.value = model.sim.variationPct;

    const varyOutputs = $("#simVaryOutputs");
    const varyTreatCosts = $("#simVaryTreatCosts");
    const varyInputCosts = $("#simVaryInputCosts");
    if (varyOutputs) varyOutputs.value = model.sim.varyOutputs ? "true" : "false";
    if (varyTreatCosts) varyTreatCosts.value = model.sim.varyTreatCosts ? "true" : "false";
    if (varyInputCosts) varyInputCosts.value = model.sim.varyInputCosts ? "true" : "false";
  }

  // =========================
  // 16) SAVE/LOAD PROJECT JSON
  // =========================
  function saveProject() {
    const payload = JSON.stringify(model, null, 2);
    const blob = new Blob([payload], { type: "application/json;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${slug(TOOL_NAME)}_${slug(model.project.name)}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  function loadProjectFile(file) {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const parsed = JSON.parse(e.target.result);
        // Basic safety: ensure correct tool name exists but do not block import
        if (parsed && typeof parsed === "object") {
          // shallow merge carefully
          Object.assign(model, parsed);
          model.meta = model.meta || {};
          model.meta.toolName = TOOL_NAME;
          model.meta.version = TOOL_VERSION;

          // Ensure raw default exists
          if (!model.raw) model.raw = {};
          if (!model.raw.fabaBeans2022) model.raw.fabaBeans2022 = DEFAULT_FABA_BEANS_RAW;

          initTreatmentDeltas();
          setToolBranding();
          bindProjectFields();
          bindSettingsFields();
          renderOutputs();
          renderTreatments();
          refreshAll();
          showToast("Project loaded.");
        }
      } catch (err) {
        console.error(err);
        alert("Failed to load project JSON: " + (err?.message || String(err)));
      }
    };
    reader.readAsText(file);
  }

  function bindProjectSaveLoad() {
    const saveBtn = $("#saveProject");
    const loadBtn = $("#loadProject");
    const loadFile = $("#loadFile");

    if (saveBtn) saveBtn.addEventListener("click", saveProject);

    if (loadBtn && loadFile) {
      loadBtn.addEventListener("click", () => {
        loadFile.value = "";
        loadFile.click();
      });
      loadFile.addEventListener("change", () => {
        const f = loadFile.files && loadFile.files[0];
        if (f) loadProjectFile(f);
      });
    }
  }

  // =========================
  // 17) REFRESH PIPELINE
  // =========================
  function refreshAll() {
    renderToplineMetrics();
    renderComparisonTable();
    renderTimeProjection();
    refreshCopilotPreview();
  }

  function bindResultsButtons() {
    const recalc = $("#recalc");
    if (recalc) {
      recalc.addEventListener("click", () => {
        refreshAll();
        showToast("Results recalculated.");
      });
    }
  }

  // =========================
  // 18) POLISH: RESULTS TAB TITLE + EXPORT FOOTER LABELS
  // =========================
  function polishUIStrings() {
    const resH2 = $("#tab-results h2");
    if (resH2) resH2.textContent = "Results (control vs treatments)";
    const introTitle = $("#tab-intro h1");
    if (introTitle) introTitle.textContent = TOOL_NAME;

    // Excel tab title
    const excelH2 = $("#tab-excel h2");
    if (excelH2) excelH2.textContent = "Excel import and export (Excel-first workflow)";

    // Copilot tab title handled in bindCopilot()

    // Header start button label
    const startBtn = $("#startBtn");
    const startBtnDup = $("#startBtn-duplicate");
    if (startBtn) startBtn.textContent = "Start with project setup";
    if (startBtnDup) startBtnDup.textContent = "Go to project setup";
  }

  // =========================
  // 19) INIT
  // =========================
  function init() {
    setToolBranding();
    polishUIStrings();

    initTabs();

    // Default scenario load
    loadDefaultScenario();

    // Bind core tabs
    bindProjectFields();
    bindSettingsFields();

    renderOutputs();
    bindOutputsButtons();

    renderTreatments();
    bindTreatmentsButtons();

    bindExportButtons();
    bindExcelButtons();
    bindCopilot();
    bindSimulation();

    bindProjectSaveLoad();
    bindResultsButtons();

    // First paint
    refreshAll();

    // Start button: jump to project
    const startBtn = $("#startBtn");
    const startBtnDup = $("#startBtn-duplicate");
    [startBtn, startBtnDup].filter(Boolean).forEach(btn => btn.addEventListener("click", () => activateTab("project")));
  }

  // Run
  document.addEventListener("DOMContentLoaded", init);
})();
