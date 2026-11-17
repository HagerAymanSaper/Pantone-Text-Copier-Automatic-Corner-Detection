# Pantone Text Copier v4

An Adobe Photoshop automation script designed to automatically place Pantone text labels inside selected pattern pieces.

Instead of manually positioning text on every piece, the script analyzes each piece's actual outline, detects suitable corners, and automatically positions a copied text label.

## Features

- **Automatic Corner Detection:** Detects convex corners formed by straight edges.
- **90° Corner Matching:** Finds corners closest to a 90-degree angle.
- **Custom Corner Selection:** Supports Bottom-left, Bottom-right, Top-left, and Top-right preferences.
- **Custom Label Distance:** Specify the distance between the text and both corner edges in centimeters.
- **Automatic Text Rotation:** Rotates labels to follow the detected corner's orientation.
- **Outline Detection:** Uses layer transparency to trace the actual shape of each pattern piece.
- **Smart Text Placement:** Checks whether the label fits inside the piece and adjusts its position when necessary.
- **Batch Processing:** Processes multiple selected pattern pieces in one operation.
- **Original Layer Protection:** Keeps the original text layer and pattern piece layers unchanged.

## Requirements

- Adobe Photoshop
- A Photoshop document containing:
  - Pattern pieces on separate layers.
  - At least one existing text layer to use as the source label.

## How to Use

1. Open your Photoshop pattern document.
2. Select only the pattern piece layers in the Layers panel using Ctrl + Click.
3. Run `Pantone_Text_Copier_v4.jsx`.
4. Select the text layer you want to copy.
5. Enter the desired distance from both corner edges (default: 0.7 cm).
6. Choose your preferred corner.
7. Click **Place Labels**.

The script will automatically analyze each selected piece and position a copy of the text inside the detected corner.

## How It Works

1. Reads the transparency of each selected layer.
2. Traces the actual outline using a Photoshop Work Path.
3. Analyzes the outline and identifies straight edges.
4. Detects convex corners and measures their interior angles.
5. Selects the corner closest to 90 degrees.
6. Positions the copied text at the requested distance from both edges.
7. Rotates the text to match the corner orientation.
8. Reports successful placements and skipped pieces.

## Important Notes

- Select pattern piece layers only, not groups.
- The source text layer must exist in the document.
- The original text and pattern piece layers are not resized or edited.
- If a label cannot fit at the requested distance, the script attempts to adjust its placement and reports the result.
- The script temporarily creates a Photoshop Work Path and removes it after processing. An existing Work Path is replaced.

## Version

**Version 4 – Automatic Corner Detection**
