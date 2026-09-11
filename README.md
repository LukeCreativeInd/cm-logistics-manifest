# CM Logistics Manifest Generator

Streamlit app for generating carrier manifests from Clean Eats and Made Active Shopify order exports.

Orders are routed by carrier tags including `CM`, `MC`, `CX`, `DK`, and `HDS`. HDS manifests use one row per carton and repeat the complete delivery details on every carton row.

HDS delivery types are `business` for Clean Eats Wholesale (`CEW`) orders and `residential` for Clean Eats Australia (`CEA`) and Made Active orders.
