{pkgs}: {
  deps = [
    pkgs.tree
    pkgs.libxcrypt
    pkgs.postgresql
    pkgs.openssl
  ];
}
