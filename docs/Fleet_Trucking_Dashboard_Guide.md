# Fleet Cost \& Maintenance Analysis Dashboard — Build Guide (v2: Trucking Fleet)

**Portfolio project for:** Asad Ullah, Data Analyst / Power BI Developer
**Tools:** Power BI Desktop, Power Query (M), DAX
**Dataset:** `Fleet\\\_Trucking\\\_Dataset\\\_2024\\\_2026.xlsx`
**Fleet:** 45 long-haul trucks, 3 depots (Dammam, Riyadh, Jeddah), Jan 2024 - Jun 2026 (30 months)

\---

## 1\. Business problem framing

A logistics company runs long-haul trucks across Saudi Arabia out of three depots.
Management wants to know:

* Which trucks cost the most to operate?
* Which trucks require frequent maintenance, and from which vendors?
* How much fuel is consumed each month, and is fuel cost rising?
* Which trucks/routes are underutilized?
* What is the total operating cost per kilometer?

This framing is what you lead with in an interview or on your portfolio page —
it's the "why" behind every visual on the dashboard.

\---

## 2\. The dataset

Open `Fleet\\\_Trucking\\\_Dataset\\\_2024\\\_2026.xlsx`. The `ReadMe` tab documents everything
below in the workbook itself.

|Sheet|Type|Grain|Rows|
|-|-|-|-|
|Calendar|Dimension|1 row per day, 2024-01-01 to 2026-06-30|912|
|FleetMaster|Dimension|1 row per truck (45 trucks: Flatbed / Refrigerated / Container / Tanker / Curtain-Side)|45|
|DriverMaster|Dimension|1 row per driver|38|
|DailyTrips|Fact|1 row per completed trip (Origin, Destination, Distance, Cargo, Trip Hours)|11,273|
|FuelLog|Fact|1 row per fuel fill-up (Liters, Fuel Price/L — cost is NOT pre-calculated)|8,851|
|MaintenanceLog|Fact|1 row per maintenance/repair event (Type, Vendor, Cost, Downtime)|1,521|

**Deliberate realism built in** (mention these in interviews — they show you think about data quality, not just visuals):

* Diesel price per liter drifts upward across the period (\~SAR 0.52-0.56/L in early 2024 to \~SAR 0.62-0.68/L by mid-2026), so a genuine fuel-cost-inflation story exists in the data.
* 3 trucks are marked `Status = Retired` in FleetMaster with no activity after retirement — good for testing filters and "active fleet" logic.
* Not every trip has a matching fuel log (\~78% do) and maintenance isn't evenly spaced — mirrors real gaps in operational data.
* Fuel cost is intentionally left as `Liters` and `FuelPricePerLiter` only, so you calculate `Fuel Cost = Liters \\\* FuelPricePerLiter` yourself in Power Query (Step 3 below) — this is a real transformation you can speak to, not just a pre-baked number.

\---

## 3\. Step-by-step build

### Step 1 — Import data

1. Power BI Desktop → **Get Data → Excel workbook** → `Fleet\\\_Trucking\\\_Dataset\\\_2024\\\_2026.xlsx`.
2. Load: `Calendar`, `FleetMaster`, `DriverMaster`, `DailyTrips`, `FuelLog`, `MaintenanceLog`. Skip `ReadMe`.

### Step 2 — Power Query transformations

**All tables**

* Confirm data types: `Date` columns as **Date**, ID columns as **Text**, `Cost`, `Liters`, `FuelPricePerLiter`, `DistanceKM`, `CargoMT`, `TripHours`, `DowntimeHours`, `CapacityMT` as **Decimal Number**.

**DailyTrips**

* Add calculated column **Route**:

```
  Route = \\\[Origin] \\\& " → " \\\& \\\[Destination]
  ```

* Remove any exact duplicate rows (Home tab → Remove Rows → Remove Duplicates) as a data-hygiene step, even if none exist — it's a habit worth demonstrating.

**FuelLog**

* Add calculated column **FuelCost**:

```
  FuelCost = \\\[Liters] \\\* \\\[FuelPricePerLiter]
  ```

**FleetMaster**

* Add calculated column **VehicleAgeYears**:

```
  VehicleAgeYears = Duration.Days(Date.From(DateTime.LocalNow()) - \\\[PurchaseDate]) / 365
  ```

**MaintenanceLog**

* Standardize `Vendor` names (Trim/Clean, fix casing) — a realistic ETL step even though this dataset is already clean; note in your case study that you'd add a fuzzy-matching step for real vendor data.

Close \& Apply.

### Step 3 — Data model (star schema)

In **Model view**, build these relationships (1 → many, single direction, dimension → fact):

* `FleetMaster\\\[VehicleID]` → `DailyTrips\\\[VehicleID]`
* `FleetMaster\\\[VehicleID]` → `FuelLog\\\[VehicleID]`
* `FleetMaster\\\[VehicleID]` → `MaintenanceLog\\\[VehicleID]`
* `DriverMaster\\\[DriverID]` → `DailyTrips\\\[DriverID]`
* `Calendar\\\[Date]` → `DailyTrips\\\[Date]`
* `Calendar\\\[Date]` → `FuelLog\\\[Date]`
* `Calendar\\\[Date]` → `MaintenanceLog\\\[Date]`

Mark `Calendar` as a **Date Table** (Table tools → Mark as Date Table → column `Date`).
This is required for the time-intelligence DAX below.

### Step 4 — DAX measures

Create a **Measures** table (`New Table → Measures = ROW("dummy", 0)`, hide the dummy column) to keep all measures organized in one place, separate from the fact tables.

**Cost**

```DAX
Total Distance = SUM(DailyTrips\\\[DistanceKM])

Fuel Cost = SUM(FuelLog\\\[FuelCost])

Maintenance Cost = SUM(MaintenanceLog\\\[Cost])

Operating Cost = \\\[Fuel Cost] + \\\[Maintenance Cost]

Cost Per KM = DIVIDE(\\\[Operating Cost], \\\[Total Distance], 0)

Avg Maintenance Cost = AVERAGE(MaintenanceLog\\\[Cost])
```

**Trips \& fleet activity**

```DAX
Trips = COUNTROWS(DailyTrips)

Active Trucks = CALCULATE(DISTINCTCOUNT(FleetMaster\\\[VehicleID]), FleetMaster\\\[Status] = "Active")

Trips per Truck = DIVIDE(\\\[Trips], DISTINCTCOUNT(DailyTrips\\\[VehicleID]), 0)
```

**Fuel efficiency**

```DAX
Total Liters = SUM(FuelLog\\\[Liters])

KM per Liter = DIVIDE(\\\[Total Distance], \\\[Total Liters], 0)
```

**Downtime \& utilization**

```DAX
Downtime Hours = SUM(MaintenanceLog\\\[DowntimeHours])

Avg Downtime per Event = DIVIDE(\\\[Downtime Hours], DISTINCTCOUNT(MaintenanceLog\\\[MaintenanceID]), 0)

Vehicle Utilization % =
DIVIDE(
    DISTINCTCOUNT(DailyTrips\\\[VehicleID]),
    CALCULATE(DISTINCTCOUNT(FleetMaster\\\[VehicleID]), ALL(DailyTrips)),
    0
)
```

**Time intelligence (needs Calendar marked as date table)**

```DAX
Operating Cost MoM % =
VAR CurrentCost = \\\[Operating Cost]
VAR PriorCost = CALCULATE(\\\[Operating Cost], DATEADD(Calendar\\\[Date], -1, MONTH))
RETURN DIVIDE(CurrentCost - PriorCost, PriorCost, 0)

Fuel Cost YTD = TOTALYTD(\\\[Fuel Cost], Calendar\\\[Date])
```

**Ranking (Top 10 expensive vehicles)**

```DAX
Vehicle Total Cost =
CALCULATE(\\\[Operating Cost], ALLEXCEPT(FleetMaster, FleetMaster\\\[VehicleID]))

Vehicle Cost Rank = RANKX(ALL(FleetMaster\\\[VehicleID]), \\\[Vehicle Total Cost], , DESC)
```

**Driver fuel efficiency (join through DailyTrips)**

```DAX
Driver Avg KM per Trip =
DIVIDE(SUM(DailyTrips\\\[DistanceKM]), COUNTROWS(DailyTrips), 0)
```

> Interview note: be ready to explain why `Vehicle Utilization %` here is defined as
> "trucks with at least one trip in the filtered period ÷ total active trucks", not a
> percentage of hours — with trip-based data (no telemetry/idle-hours feed) that's the
> honest measure you can actually calculate, and saying so shows you understand your
> data's limits rather than fabricating precision.

### Step 5 — Dashboard pages

**Page 1 — Executive Overview**

* KPI cards (top row): Total Cost, Fuel Cost, Maintenance Cost, Trips, Active Trucks, KM/L.
* Line chart: Fuel Cost + Maintenance Cost trend by month (`Calendar\\\[MonthYear]` on axis).
* Column chart: Maintenance Cost by Month.
* Bar chart: Vehicle Utilization % or Trips by Vehicle Type.
* Table: Top 10 Expensive Vehicles (filter `Vehicle Cost Rank ≤ 10`).
* Donut/bar: Maintenance Type breakdown (Oil Change, Brake Repair, Engine Repair, etc.).
* Bar: Cost by Vehicle Type (Flatbed / Refrigerated / Container / Tanker / Curtain-Side).

**Page 2 — Vehicle Performance**

* Cards: Total KM, KM/L, Total Trips (scoped to selected truck via slicer).
* Distance by Vehicle (bar).
* Fuel Consumption trend (line).
* Driver Performance (table: driver, trips, avg KM/trip).
* Top Routes (table or map: Origin → Destination pairs, trip count, avg distance).
* Vehicle Utilization % by truck.

**Page 3 — Maintenance Analysis**

* Maintenance Type (bar).
* Maintenance Vendor (bar) — which vendors are used most / cost most.
* Monthly Maintenance cost trend.
* Downtime Trend (line, by month).
* Repair Frequency by vehicle (table, sorted descending).

**Page 4 — Cost Analysis**

* Fuel vs Maintenance Cost (combo chart).
* Cost per Vehicle (bar, sorted).
* Cost per KM trend (line, by month).
* Monthly Expenses (column).
* Operating Cost trend with MoM % (line + labels).

**Page 5 — Fleet Overview**

* Map: Origin depot → Destination (use `DailyTrips\\\[Origin]`/`\\\[Destination]` with Bing map coordinates, or a static route table).
* KPI cards: Active Trucks, Under Maintenance, Retired, Avg Trip Distance, Avg KM/L.

### Step 6 — Slicers (add to every page or a shared filter pane)

Date, Vehicle (`FleetMaster\\\[VehicleID]`), Driver, Vehicle Type, Vendor (`MaintenanceLog\\\[Vendor]`), Maintenance Type, Status.

### Step 7 — Advanced features worth adding (these are what separate a portfolio piece from a homework exercise)

* **Drill-through** from Page 1's "Top 10 Expensive Vehicles" table to a Vehicle Details page.
* **Tooltip page**: hovering a bar/line shows a mini card with vehicle make/model/depot.
* **Bookmarks**: toggle between "Fleet View" and "Maintenance View" with a button.
* **Dynamic titles**: title changes based on the selected slicer value (e.g. "Operating Cost — Dammam Depot").
* **Field parameters**: let the user switch the Page-1 line chart between Cost / Distance / Fuel with one control.
* **Conditional formatting**: red/amber/green on the Cost Per KM and Downtime cards.
* **Navigation buttons** between the 5 pages instead of relying on the page tabs.

### Step 8 — Formatting checklist

* One consistent theme (View → Themes) instead of per-visual styling.
* Color logic: blue = fuel, amber/red = maintenance \& downtime (severity), green = utilization/efficiency.
* Turn off gridlines/legends where a chart title already makes the meaning clear.
* Footer with your name and "Portfolio project — synthetic data" so it's clearly labeled as a demo, not a real client's numbers.

### Step 9 — Package it for your portfolio

1. Save as `.pbix`.
2. Export a 2-3 page PDF snapshot (File → Export → Export to PDF).
3. Screenshot each page at 1920×1080.
4. Write a short case-study paragraph (business problem → data model → key insight → visuals) to sit next to the screenshots on LinkedIn/portfolio site.
5. If you want it on GitHub: include the dataset, the `.pbix`, screenshots, and a README describing the schema and DAX — reuse the ReadMe tab content from the workbook as a starting point.

\---

## 4\. Resume / portfolio bullet points

Pick 2-3, don't stack all of them:

* *Designed and developed an interactive Power BI dashboard to monitor fleet maintenance expenses, fuel consumption, vehicle utilization, operating costs, and overall performance across a 45-truck fleet over a 30-month period.*
* *Built a star-schema data model integrating fleet, trip, fuel, maintenance, and driver data (20,000+ combined rows) using Power Query, including a manual Fuel Cost calculation to mirror a real unprocessed source feed.*
* *Created 20+ DAX measures and KPIs — operating cost, cost per kilometer, fuel efficiency (KM/L), vehicle utilization, and RANKX-based cost ranking — to identify the highest-cost vehicles and rising maintenance trends.*
* *Implemented drill-through pages, dynamic slicers, tooltips, bookmarks, and conditional formatting to improve operational visibility and support data-driven fleet decisions.*

## 5\. Likely interview questions

* **"Why is Fuel Cost calculated in Power Query instead of coming from the source?"** → Mirrors how a real fuel-card/telematics export would arrive (liters + price, not a pre-computed total) — shows you handle raw source data, not just pre-cleaned demo files.
* **"How did you define utilization without telemetry/idle-hours data?"** → Explain the trip-based definition above, and what you'd add with a telematics feed (idle hours, engine-on time).
* **"Why three depots and route-based trips instead of random data?"** → Realistic Saudi road distances (Dammam-Riyadh 400km, Riyadh-Jeddah 950km, etc.) so cost-per-KM and route comparisons are analytically meaningful, not arbitrary.
* **"What would you change for production use?"** → Incremental refresh on DailyTrips/FuelLog (they grow daily), row-level security by depot, and a data-quality page flagging trucks with no fuel log for 14+ days.

\---

*This dataset is synthetic — generated for portfolio use — but structured to mirror a real Saudi trucking fleet's systems, consistent with the Fleet Cost, Accounts Payable, and Inventory reports you've built at Square Root Logistics.*

