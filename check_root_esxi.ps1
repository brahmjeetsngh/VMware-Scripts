get-vmhost | %{
  $null = connect-viserver $_.name `
  -user root -password "enter password" -EA 0

  if (-not ($?)) {
    write-warning "Password failed for $($_.name)"
  } else {
    Disconnect-VIServer $_.name -force -confirm:$false
  }
}