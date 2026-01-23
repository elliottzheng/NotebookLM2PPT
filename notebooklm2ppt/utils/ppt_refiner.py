import json
import numpy as np
from PIL import Image, ImageDraw
import os
import requests
from sklearn.cluster import DBSCAN
from spire.presentation.common import *
from spire.presentation import *
from .ppt_combiner import clean_ppt

def recursive_blocks(blocks):
    result = []
    for block in blocks:
        if "blocks" in block:
            result.extend(recursive_blocks(block["blocks"]))
        else:
            result.append(block)
    return result


def get_scaled_para_blocks(resize_scale, pdf_info, page_index, cond = 'no_image'):
    para_blocks = pdf_info[page_index]['para_blocks'] + pdf_info[page_index]['discarded_blocks']
    para_blocks = recursive_blocks(para_blocks)

    scaled_para_blocks = []
    for block in para_blocks:
        if cond == 'no_image' and block['type'] in ['image_body',"table_body"]:
            continue
        if cond == 'only_image' and block['type'] not in ['image_body',"table_body"]:
            continue
        scaled_block = block.copy()
        bbox = block['bbox']
        scaled_bbox = [
            bbox[0] * resize_scale,
            bbox[1] * resize_scale,
            bbox[2] * resize_scale,
            bbox[3] * resize_scale
        ]
        scaled_block['bbox'] = scaled_bbox
        scaled_para_blocks.append(scaled_block)
    return scaled_para_blocks



def compute_iou(boxA, boxB):
    # box = [left, top, right, bottom]
    xA = max(boxA[0], boxB[0])
    yA = max(boxA[1], boxB[1])
    xB = min(boxA[2], boxB[2])
    yB = min(boxA[3], boxB[3])

    interWidth = max(0, xB - xA)
    interHeight = max(0, yB - yA)
    interArea = interWidth * interHeight

    boxAArea = (boxA[2] - boxA[0]) * (boxA[3] - boxA[1])
    boxBArea = (boxB[2] - boxB[0]) * (boxB[3] - boxB[1])

    iou = interArea / float(boxAArea + boxBArea - interArea)

    return iou

def compute_ious(left, top, height, width, scaled_para_blocks):
    bbox = [left, top, left + width, top + height]
    ious = []
    for block in scaled_para_blocks:
        block_bbox = block['bbox']
        iou = compute_iou(bbox, block_bbox)
        ious.append(iou)
    return ious

def download_image(image_url, tmp_image_path):
    if os.path.exists(tmp_image_path):
        return
    response = requests.get(image_url)

    with open(tmp_image_path, 'wb') as f:
        f.write(response.content)

def load_json(file_path):
    with open(file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data




def compute_edge_diversity(image_cv, left, top, right, bottom):
    left, top, right, bottom = round(left), round(top), round(right), round(bottom)
    top_edge = image_cv[top-1:top, left:right]
    bottom_edge = image_cv[bottom:bottom+1, left:right]
    left_edge = image_cv[top:bottom, left-1:left]
    right_edge = image_cv[top:bottom, right:right+1]
    edges = [top_edge, bottom_edge, left_edge, right_edge]
    flatten_points = np.concatenate([edge.reshape(-1, 3) for edge in edges], axis=0)
    clustering = DBSCAN(eps=2, min_samples=3).fit(flatten_points)

    counts = []
    for label in np.unique(clustering.labels_):
        count = np.sum(clustering.labels_==label)
        counts.append(count)

    main_ratio = np.max(counts) / np.sum(counts)  # 主要颜色占比

    main_color = flatten_points[clustering.labels_==np.argmax(counts)].mean(axis=0)
    return 1 - main_ratio, main_color



def get_indices_from_png_names(png_names):
    indices = []
    for name in png_names:
        base = os.path.basename(name)
        index_str = base.replace('page_', '').replace('.png', '')
        indices.append(int(index_str) - 1)
    return indices


def refine_ppt(tmp_image_dir, json_file, ppt_file, png_dir, png_files, final_out_ppt_file):
    png_files = [os.path.join(png_dir, name) for name in png_files]
    indices = get_indices_from_png_names(png_files)
    os.makedirs(tmp_image_dir, exist_ok=True)
    data = load_json(json_file)
    pdf_info = data['pdf_info']

    pdf_info = [pdf_info[i] for i in indices] # 只保留需要的页码信息

    pdf_w, _ = pdf_info[0]['page_size']
    

    presentation = Presentation()
    presentation.LoadFromFile(ppt_file)

    ppt_H, ppt_W = presentation.SlideSize.Size.Height, presentation.SlideSize.Size.Width

    ppt_scale = ppt_W / pdf_w

    assert len(png_files) == len(pdf_info) == len(presentation.Slides)
    
    for page_index, slide in enumerate(presentation.Slides):
        print(f"优化 第 {page_index+1}/{len(png_files)} 页...")
        scaled_para_blocks = get_scaled_para_blocks(ppt_scale,pdf_info, page_index)
        # 删除不相关文本框, 统一字体
        for i in range(slide.Shapes.Count - 1, -1, -1):
            # Check if those shapes are images
            shape = slide.Shapes[i]
            print("---")
            if "IAutoShape" not in str(type(shape)):
                slide.Shapes.RemoveAt(i) # 删除非文本框形状
                continue
            # Get the first paragraph of the shape
            paragraph = shape.TextFrame.Paragraphs[0]        

            left, top, text, width, height = shape.Left,shape.Top, shape.TextFrame.Text,shape.Width,shape.Height
            print(f"text:{text} left:{left} top:{top} width:{width} height:{height}")
            ious = compute_ious(left, top, height, width, scaled_para_blocks)
            # print(len(ious))

            if np.max(ious)>0.01:
                print("max iou:",np.max(ious))
                
                neareast_block = scaled_para_blocks[np.argmax(ious)]
                if neareast_block['type'] in ['title','text']:
                    print(neareast_block)
            else:
                print("invalid")
                slide.Shapes.RemoveAt(i)
                continue


            assert left+width <= ppt_W +10
            assert top+height <= ppt_H +10

            # Create a font
            newFont = TextFont("微软雅黑")

            # Loop through the text ranges in the paragraph
            for textRange in paragraph.TextRanges:
                textRange.LatinFont = newFont # 更换字体

        # 替换图片    
        image_blocks = get_scaled_para_blocks(ppt_scale,pdf_info, page_index,'only_image')
        for image_block in image_blocks:
            for line in image_block['lines']:
                for span in line['spans']:
                    tmp_image_path = os.path.join(tmp_image_dir, os.path.basename(span['image_path']))

                    download_image(span['image_path'], tmp_image_path)

                    left, top, right, bottom = image_block['bbox']

                    delta_y = 2 # 下移两个像素

                    rect1 = RectangleF.FromLTRB(left, top + delta_y, right, bottom + delta_y)
                    image = slide.Shapes.AppendEmbedImageByPath(ShapeType.Rectangle, tmp_image_path, rect1)
                    image.Line.FillType = FillFormatType.none
                    image.ZOrderPosition = 0  # 设置图片在最底层
        

        # 替换背景    
        background = slide.SlideBackground
        old_bg_file = "old_bg.png"
        try:
            background.Fill.PictureFill.Picture.EmbedImage.Image.Save(old_bg_file)
            old_bg_cv = np.array(Image.open(old_bg_file))
            os.remove(old_bg_file)
        except:
            print("No existing background image found in slide ", page_index)
            old_bg_cv = None
        # 替换背景    
        background.Type = BackgroundType.Custom

        # Set the fill mode of the slide's background as a picture fill
        background.Fill.FillType = FillFormatType.Picture

        # Add an image to the image collection of the presentation

        png_file = png_files[page_index]
        image_cv = Image.open(png_file)
        image_cv = np.array(image_cv)

        image_h, image_w, _ = image_cv.shape

        if old_bg_cv is not None:
            # 使用PIL进行resize
            old_bg_pil = Image.fromarray(old_bg_cv)
            old_bg_pil = old_bg_pil.resize((image_w, image_h), Image.BICUBIC)
            old_bg_cv = np.array(old_bg_pil)

        image_scale = image_w / pdf_w

        text_blocks = get_scaled_para_blocks(image_scale, pdf_info, page_index, cond='no_image')


        for text_block in text_blocks:
            bbox = text_block['bbox']
            l, t, r, b = map(round, bbox)
            diversity, fill_color = compute_edge_diversity(image_cv, l, t, r, b)
            fill_color = fill_color.astype(np.uint8).tolist()
            print("div=", diversity, " text_block=", text_block)
            if old_bg_cv is None or diversity < 0.5: # 边缘多样性低，认为是纯色区域，则可以直接填充
                image_cv[t:b, l:r] = fill_color
            else: # 边缘多样性高，保留原背景
                image_cv[t:b, l:r] = old_bg_cv[t:b, l:r] # 保留原背景的前提是要有原背景图

        tmp_bg_file = png_file.replace('.png', '_bg.png')
        Image.fromarray(image_cv).save(tmp_bg_file)
        stream = Stream(tmp_bg_file)

        imageData = presentation.Images.AppendStream(stream)
        # Set the image as the slide's background
        background.Fill.PictureFill.FillType = PictureFillType.Stretch
        background.Fill.PictureFill.Picture.EmbedImage = imageData
        
    presentation.SaveToFile(final_out_ppt_file, FileFormat.Pptx2019)

    print(f"优化完成! 输出文件: {final_out_ppt_file}")
    clean_ppt(final_out_ppt_file,final_out_ppt_file)


