{
  description = "HTML slide decks → PowerPoint (slides-to-pptx skill)";

  inputs = {
    nixpkgs.url = "github:NixOS/nixpkgs/nixpkgs-unstable";
    utils.url = "github:numtide/flake-utils";
  };

  outputs = { self, nixpkgs, utils }: utils.lib.eachDefaultSystem (system:
    let
      pkgs = nixpkgs.legacyPackages.${system};
    in
    {
      devShells.default = pkgs.mkShell {
        buildInputs = with pkgs; [
          # html2pptx.mjs runtime + npm for playwright/pptxgenjs/jszip.
          # Everything else the skill needs ships in those npm packages: the .pptx is
          # written by pptxgenjs, and both capture and preview run in Playwright's
          # own Chromium. No office suite is involved.
          nodejs_22
          # inspecting the OOXML package by hand
          unzip
          libxml2
        ];
      };
      # Kept for `nix-shell`/older callers.
      devShell = self.devShells.${system}.default;
    }
  );
}
