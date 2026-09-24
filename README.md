# EXCEL_DATA-EXPLORATION_ASSIGNMENT-1

## Assignment Overview

The purpose of this assignment is to practice important Excel functions used for:

- Basic calculations
- Counting records
- Finding minimum and maximum values
- Conditional logic
- Conditional totals and counts
- Text extraction from Product IDs

---

## 1. SUM, COUNT, AND AVERAGE

### Question 1: What is the total price of all products in the dataset?

The `SUM` function is used to add all product prices.

```excel
=SUM(D2:D35)
```

### Explanation

- `D2:D35` represents all values in the **Price** column.
- `SUM` adds all prices in that range.

**Result: 10100**

---

### Question 2: How many products are there in the dataset?

The `COUNT` function is used to count numeric values.

```excel
=COUNT(D2:D35)
```

### Explanation

- The Price column contains one numeric price for each product.
- Therefore, counting the price cells gives the total number of products.

**Result: 34**

---

### Question 3: Calculate the average price of the products

The `AVERAGE` function is used to calculate the average product price.

```excel
=AVERAGE(D2:D35)
```

### Explanation

The average is calculated as:

```text
Average = Total of all prices / Number of products
```

The original result is approximately:

```text
297.0588235
```

To display only two decimal places:

1. Select the result cell.
2. Go to **Home > Number**.
3. Click **Decrease Decimal** until only two decimal places remain.

**Result: 297.06**

---

## 2. MINIMUM AND MAXIMUM

### Question 4: Determine the minimum price among all products

The `MIN` function returns the smallest value in a range.

```excel
=MIN(D2:D35)
```

### Explanation

Excel checks all prices from `D2` to `D35` and returns the lowest value.

**Result: 30**

---

### Question 5: Find the maximum price among all products

The `MAX` function returns the largest value in a range.

```excel
=MAX(D2:D35)
```

### Explanation

Excel checks all prices from `D2` to `D35` and returns the highest value.

**Result: 1000**

---

## 3. IF FUNCTION

### Question 6: Create a new column named Price Range

The requirement is to categorize products based on their price.

### Condition

- Price greater than or equal to `$500` → **High Price**
- Price less than `$500` → **Standard Price**

The `IF` function checks whether a condition is true or false.

### Syntax

```excel
=IF(condition,value_if_true,value_if_false)
```

### Formula

```excel
=IF(D2>=500,"High Price","Standard Price")
```

### Explanation

- `D2>=500` checks whether the product price is greater than or equal to 500.
- If the condition is true, Excel returns `High Price`.
- If the condition is false, Excel returns `Standard Price`.

### Example

| Price | Price Range |
| ---: | --- |
| 1000 | High Price |
| 80 | Standard Price |
| 500 | High Price |
| 400 | Standard Price |

After entering the formula in the first row, drag it down to apply the formula to the remaining rows.

---

## 4. SUMIF AND COUNTIF

### Question 7: Calculate the total price of products in the Electronics category

The `SUMIF` function adds values only when a condition is satisfied.

### Syntax

```excel
=SUMIF(range,criteria,sum_range)
```

### Formula

```excel
=SUMIF(F2:F35,"Electronics",D2:D35)
```

### Explanation

```text
F2:F35        -> Checks the Category column
"Electronics" -> Required condition
D2:D35        -> Adds the matching product prices
```

Excel checks each row in the Category column.

If the category is `Electronics`, the corresponding value from the Price column is added.

**Result: 8050**

---

### Question 8: Determine the count of products with a price less than $100

The `COUNTIF` function counts cells that satisfy a given condition.

### Syntax

```excel
=COUNTIF(range,criteria)
```

### Formula

```excel
=COUNTIF(D2:D35,"<100")
```

### Explanation

- `D2:D35` represents the Price column.
- `"<100"` tells Excel to count only prices below 100.

The price `100` is not counted because the condition is:

```text
<100
```

This means strictly less than 100.

If the requirement were less than or equal to 100, the formula would be:

```excel
=COUNTIF(D2:D35,"<=100")
```

**Result: 11**

---

## 5. TEXT FUNCTIONS - LEFT, RIGHT, AND MID

The Product ID follows a structure like:

```text
28-JAN-US
```

The Product ID contains:

```text
28  -> Day
JAN -> Month
US  -> Country Code
```

---

### Question 9: Create a new column named Day using the LEFT function

The `LEFT` function returns characters from the beginning of a text string.

### Syntax

```excel
=LEFT(text,num_chars)
```

### Formula

```excel
=LEFT(A2,2)
```

### Explanation

- `A2` contains the Product ID.
- `2` means Excel should extract the first two characters.

### Example

Product ID:

```text
28-JAN-US
```

Formula:

```excel
=LEFT(A2,2)
```

Result:

```text
28
```

Another example:

```text
03-MAR-US
```

Result:

```text
03
```

After entering the formula, drag it down to apply it to all Product IDs.

---

### Question 10: Create a new column named Country Code using the RIGHT function

The `RIGHT` function returns characters from the end of a text string.

### Syntax

```excel
=RIGHT(text,num_chars)
```

### Formula

```excel
=RIGHT(A2,2)
```

### Explanation

- `A2` contains the Product ID.
- `2` means Excel should extract the last two characters.

### Example

Product ID:

```text
28-JAN-US
```

Formula:

```excel
=RIGHT(A2,2)
```

Result:

```text
US
```

Another example:

```text
17-JUN-IN
```

Result:

```text
IN
```

After entering the formula, drag it down to apply it to all Product IDs.

---

### Question 11: Create a new column named Month using the MID function

The `MID` function extracts characters from the middle of a text string.

### Syntax

```excel
=MID(text,start_num,num_chars)
```

### Formula

```excel
=MID(A2,4,3)
```

### Explanation

For the Product ID:

```text
28-JAN-US
```

Character positions are:

```text
1 = 2
2 = 8
3 = -
4 = J
5 = A
6 = N
7 = -
8 = U
9 = S
```

The formula:

```excel
=MID(A2,4,3)
```

means:

- Start from the 4th character.
- Extract 3 characters.

Result:

```text
JAN
```

Another example:

```text
15-FEB-US
```

Result:

```text
FEB
```

After entering the formula, drag it down to apply it to all Product IDs.

---

## Final Results

| Task | Result |
| --- | ---: |
| Total price of all products | 10100 |
| Total number of products | 34 |
| Average product price | 297.06 |
| Minimum product price | 30 |
| Maximum product price | 1000 |
| Total price of Electronics products | 8050 |
| Products with price less than $100 | 11 |

---

## Excel Functions Used

- `SUM`
- `COUNT`
- `AVERAGE`
- `MIN`
- `MAX`
- `IF`
- `SUMIF`
- `COUNTIF`
- `LEFT`
- `RIGHT`
- `MID`

---

