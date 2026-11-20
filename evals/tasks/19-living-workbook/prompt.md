sales-book.xlsx is our sales workbook (Inputs, Products, Orders, Summary and Insights). Please make these 20 changes, in this order. Later ones build on earlier ones, so each starts from the workbook the earlier ones leave. Keep everything else as it is.

1. Inputs: set the tax rate to 22%.
2. Products: raise Gadget's price to 34.00.
3. Products: add a product at the bottom of the table: Gizmo Pro, Hardware, price 48.00, cost 21.00.
4. Orders: add an order at the bottom of the table: 2026-03-18, North, Gizmo Pro, 12 units.
5. Orders: the 2026-01-14 order (East, Suite) was for 4 units, not 3. Correct it.
6. Orders: the 2026-01-19 order (Support) was South's, not West's. Correct it.
7. Inputs: set the bulk threshold to 8 units.
8. Summary: add an "Orders" column in G: the header in G3, the number of orders of each region in G4:G7 (counted from the Orders table), and the total in G8.
9. Summary: in A21 put "Average order value", and in B21 the total revenue divided by the number of orders (B8/G8).
10. Rename the Insights sheet to Rankings (Summary!B20 reads its first cell).
11. Products: Toolkit's category is Service, not Software.
12. Orders: delete the 2026-01-27 order (West, Gizmo); it was cancelled.
13. Add a sheet called Targets at the end, with Region and Target in A1:B1, then North 1500, South 800, East 650 and West 900 in A2:B5.
14. Summary: add a "Gap to target" column in H: the header in H3, each region's revenue minus its target (looked up in Targets by region name) in H4:H7, and the total in H8.
15. Inputs: set the target margin to 55%.
16. Summary: show Margin % (F4:F8) with one decimal, as 0.0%.
17. Products: sort the table by price, highest first.
18. Inputs: add "Shipping per order" in A6 with 4.50 in B6, and name that cell ShipPerOrder.
19. Orders: add a column "Shipping" to the table, in column I. Each order costs ShipPerOrder to ship, except bulk orders (units at or over the bulk threshold), which ship free. Use the names, not the numbers.
20. Summary: in A22 put "Shipping cost" and in B22 the total of the Shipping column; in A23 put "Margin after shipping" and in B23 the total Margin (D8) less the shipping cost.

When you are done, tell me the total revenue and the margin after shipping.
