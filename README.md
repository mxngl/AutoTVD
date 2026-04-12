# AutoTVD Tool
**Island Team 2026 @ Stanford AEC GLobal Teamwork**

Reads Revit QTO schedule CSVs from a GitHub repository, applies cost mapping,
and generates an HTML dashboard for Target Value Design review.

**Cluster logic:**
- Clusters A–C (Substructure / Shell / Interiors):
        quantities come from the QTO takeoff.
- All other clusters (Services, Equipment, Sitework …):
        quantities come from the "Fixed Quantity" column in cost_data.csv only since the other cluster's elements are currently not represented correctly inside of Revit.
        If a row has no Fixed Quantity it contributes $0.
