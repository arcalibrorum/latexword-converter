local equation_number = 0

function Math(el)
  if el.mathtype == "DisplayMath" then
    equation_number = equation_number + 1
    return {
      el,
      pandoc.Space(),
      pandoc.Str("(" .. equation_number .. ")")
    }
  end
end
