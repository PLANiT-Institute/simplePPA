# PPA Cost Calculation Validation Results

## Overview
This document validates the cost calculation logic in `libs/calculator.py` using simple test cases with a constant demand of 1 MW.

## Test Setup
- **Load Capacity**: 1 MW (constant at 100% for all 24 hours)
- **Total Daily Energy**: 24 MWh = 24,000 kWh
- **Grid Rate**: 150 KRW/kWh (constant)
- **Grid Contract Fee**: 10,000 KRW/kW
- **PPA Price**: 170 KRW/kWh
- **Solar Pattern**: Realistic daily pattern with peak at noon (5 MWh total generation at 100% capacity)

## Test Results

### ✅ Test 1: Grid Only (0% PPA Coverage)
**Scenario**: No PPA, all energy from grid

**Expected Calculation**:
- Energy Cost: 24,000 kWh × 150 KRW/kWh = 3,600,000 KRW
- Demand Charge: 1,000 kW × 10,000 KRW/kW = 10,000,000 KRW
- **Total**: 13,600,000 KRW

**Actual Result**: 13,600,000 KRW ✅

**Cost per kWh**: 566.67 KRW/kWh
- 150 KRW/kWh (energy) + 416.67 KRW/kWh (demand charge)

---

### ✅ Test 2: 100% PPA Coverage
**Scenario**: PPA capacity equals load capacity (1 MW)

**Expected Calculation**:
- PPA Generation: 5,000 kWh (20.8% of total load)
- PPA Cost: 5,000 kWh × 170 KRW/kWh = 850,000 KRW
- Remaining Load: 19,000 kWh from grid
- Grid Energy Cost: 19,000 kWh × 150 KRW/kWh = 2,850,000 KRW
- Grid Demand Charge: 1,000 kW × 10,000 KRW/kW = 10,000,000 KRW
- **Total**: 13,700,000 KRW

**Actual Result**: 13,700,000 KRW ✅

**Analysis**:
- PPA is MORE expensive than grid only (-100,000 KRW worse)
- This makes sense because:
  - PPA price (170 KRW/kWh) > Grid energy rate (150 KRW/kWh)
  - Solar only covers 20.8% of load
  - Still pay full demand charge since grid is needed 24/7

**Cost per kWh**: 570.83 KRW/kWh

---

### ✅ Test 3: 200% PPA Coverage (with Reselling)
**Scenario**: PPA capacity is 2× load capacity, with reselling enabled at 90% of PPA price

**Expected Calculation**:
- PPA Generation: 10,000 kWh (with 100% take-or-pay must purchase all)
- PPA Purchase Cost: 10,000 kWh × 170 KRW/kWh = 1,700,000 KRW
- **Hourly Mismatch Effects**:
  - During peak solar hours: PPA > load → 2,600 kWh excess resold
  - During other hours: PPA < load → 16,600 kWh from grid
- Resale Revenue: 2,600 kWh × 170 × 0.9 = 397,800 KRW
- Net PPA Cost: 1,700,000 - 397,800 = 1,302,200 KRW
- Grid Energy Cost: 16,600 kWh × 150 KRW/kWh = 2,490,000 KRW
- Grid Demand Charge: 1,000 kW × 10,000 KRW/kW = 10,000,000 KRW
- **Total**: 13,792,200 KRW

**Actual Result**: 13,792,200 KRW ✅

**Key Insight**:
The hour-by-hour calculation is critical! Even with 10,000 kWh PPA generation vs 24,000 kWh total load:
- You can't simply subtract 10,000 from 24,000 to get grid usage
- Solar peaks midday when you still have full demand
- Creates both excess (resold) and deficit (grid purchase) in different hours

---

## Key Findings

### 1. **Cost Calculation is Accurate**
All three test scenarios match expected values exactly, confirming the calculation logic in `libs/calculator.py` is correct.

### 2. **Hour-by-Hour Logic is Essential**
The calculator properly handles the temporal mismatch between solar generation and load demand:
- Can't aggregate daily totals and assume they offset
- Must track excess and deficit hour-by-hour

### 3. **Demand Charge Dominates Cost**
For this scenario:
- Demand charge: 10,000,000 KRW (73.5% of total cost)
- Energy cost: ~3,600,000 KRW (26.5% of total cost)

This explains why:
- Simply using cheaper PPA doesn't save much money
- Reducing peak grid demand is critical
- ESS can be valuable for demand charge reduction

### 4. **PPA vs Grid Economics**
In these tests:
- PPA at 170 KRW/kWh is MORE expensive than grid energy at 150 KRW/kWh
- But real-world grids have time-varying rates (peak/off-peak)
- PPA is most valuable during high-rate periods

### 5. **Cost per kWh Includes Demand Charge**
- Grid only: 566.67 KRW/kWh (total cost ÷ energy consumed)
- This is 3.8× the energy rate due to demand charge
- Important for evaluating true energy costs

## Validation Script Usage

Run the validation:
```bash
python validate_cost.py
```

The script tests three scenarios and validates that calculated costs match expected values based on manual calculations.

## Recommendations

1. **For Analysis**: Always consider demand charges in addition to energy costs
2. **For Optimization**: Focus on reducing peak grid demand, not just energy costs
3. **For Comparison**: Use total cost per kWh, not just energy rates
4. **For PPA Sizing**: Consider temporal alignment between solar and load patterns

## Conclusion

✅ **All validation tests pass**. The cost calculation logic correctly handles:
- Multiple energy sources (PPA + Grid)
- Hour-by-hour temporal matching
- Reselling excess energy
- Demand charge calculation
- Cost aggregation

The calculations are mathematically sound and match expected results for simple test cases.
