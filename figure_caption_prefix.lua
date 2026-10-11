local figure_number = 0

local function caption_already_numbered(caption)
  local text = pandoc.utils.stringify(caption)
  return text:match("^%s*[Ff]ig%.") or text:match("^%s*[Ff]igure%s+%d")
end

local function add_figure_prefix(el)
  if el.caption and #el.caption > 0 and not caption_already_numbered(el.caption) then
    figure_number = figure_number + 1
    table.insert(el.caption, 1, pandoc.Space())
    table.insert(el.caption, 1, pandoc.Str("Fig. " .. figure_number .. "."))
    return el
  end
end

function Image(img)
  return add_figure_prefix(img)
end

function Figure(fig)
  return add_figure_prefix(fig)
end
