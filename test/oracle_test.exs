defmodule OracleTest do
  use ExUnit.Case

  alias Elixlsx.Workbook
  alias Elixlsx.Sheet

  @oracle XlsxReader

  defp check(bin) do
    {:ok, package} = @oracle.open(bin, source: :binary)
    {:ok, _sheets} = @oracle.sheets(package)

    :ok
  end

  test "empty workbook" do
    assert {:ok, {_, bin}} = Elixlsx.write_to_memory(%Workbook{}, "")
    assert check(bin)
  end

  test "empty sheet" do
    s = Sheet.with_name("1")
    wk = %Workbook{} |> Workbook.append_sheet(s)

    assert {:ok, {_, bin}} = Elixlsx.write_to_memory(wk, "")
    assert check(bin)
  end

  test "multiple sheets" do
    s1 = Sheet.with_name("1")
    s2 = Sheet.with_name("2")
    wk = %Workbook{} |> Workbook.append_sheet(s1) |> Workbook.append_sheet(s2)

    assert {:ok, {_, bin}} = Elixlsx.write_to_memory(wk, "")
    assert check(bin)
  end

  test "date cells" do
    s = Sheet.with_name("1") |> Sheet.set_cell("A4", {{2015, 11, 30}, {21, 20, 38}}, yyyymmdd: true)

    wk = %Workbook{} |> Workbook.append_sheet(s)

    assert {:ok, {_, bin}} = Elixlsx.write_to_memory(wk, "")

    assert check(bin)
  end

  test "empty date cells" do
    s = Sheet.with_name("1") |> Sheet.set_cell("A4", "", yyyymmdd: true)

    wk = %Workbook{} |> Workbook.append_sheet(s)

    assert {:ok, {_, bin}} = Elixlsx.write_to_memory(wk, "")

    assert check(bin)
  end
end
